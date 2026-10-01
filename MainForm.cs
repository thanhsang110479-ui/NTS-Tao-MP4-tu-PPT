using System.Diagnostics;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace NTS.TaoMP4;

public sealed class MainForm : Form
{
    private readonly TextBox pptBox = new() { ReadOnly = true, Dock = DockStyle.Fill };
    private readonly TextBox zipBox = new() { ReadOnly = true, Dock = DockStyle.Fill };
    private readonly Label status = new() { AutoSize = true, Text = "Chưa chọn dữ liệu." };
    private readonly ProgressBar progress = new() { Dock = DockStyle.Fill, Minimum = 0, Maximum = 100 };
    private readonly Button makeButton = new() { Text = "TẠO VIDEO MP4", Height = 46, Enabled = false, Dock = DockStyle.Fill };
    private string? pptPath, zipPath;

    public MainForm()
    {
        Text = "NTS – Tạo MP4 từ PPT";
        Width = 760; Height = 520; StartPosition = FormStartPosition.CenterScreen;
        Font = new Font("Segoe UI", 10);
        var layout = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(24), ColumnCount = 3, RowCount = 10 };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120));

        var logo = new PictureBox { Width = 88, Height = 88, SizeMode = PictureBoxSizeMode.Zoom, Anchor = AnchorStyles.None };
        var lp = Path.Combine(AppContext.BaseDirectory, "Assets", "logo.png");
        if (File.Exists(lp)) logo.Image = Image.FromFile(lp);

        var title = new Label { Text = "NTS – TẠO MP4 TỪ PPT", AutoSize = true, Font = new Font("Segoe UI", 20, FontStyle.Bold), Anchor = AnchorStyles.Left };
        var author = new Label { Text = "Tác giả: Nguyễn Thanh Sang", AutoSize = true, Font = new Font("Segoe UI", 11), Anchor = AnchorStyles.Left };

        layout.Controls.Add(logo, 0, 0); layout.SetRowSpan(logo, 2);
        layout.Controls.Add(title, 1, 0); layout.SetColumnSpan(title, 2);
        layout.Controls.Add(author, 1, 1); layout.SetColumnSpan(author, 2);

        var b1 = new Button { Text = "Chọn PPT/PPTX", Dock = DockStyle.Fill };
        var b2 = new Button { Text = "Chọn ZIP MP3", Dock = DockStyle.Fill };
        layout.Controls.Add(new Label { Text = "PowerPoint:", Anchor = AnchorStyles.Left, AutoSize = true }, 0, 3);
        layout.Controls.Add(pptBox, 1, 3); layout.Controls.Add(b1, 2, 3);
        layout.Controls.Add(new Label { Text = "Gói âm thanh:", Anchor = AnchorStyles.Left, AutoSize = true }, 0, 4);
        layout.Controls.Add(zipBox, 1, 4); layout.Controls.Add(b2, 2, 4);
        layout.Controls.Add(status, 0, 6); layout.SetColumnSpan(status, 3);
        layout.Controls.Add(progress, 0, 7); layout.SetColumnSpan(progress, 3);
        layout.Controls.Add(makeButton, 0, 8); layout.SetColumnSpan(makeButton, 3);
        layout.Controls.Add(new Label { Text = "V1: tạo video từ hình tĩnh của từng slide; không phát Animation/Transition. File gốc không bị thay đổi.", AutoSize = true }, 0, 9);
        layout.SetColumnSpan(layout.GetControlFromPosition(0,9), 3);

        Controls.Add(layout);
        b1.Click += (_,__) => PickPpt();
        b2.Click += (_,__) => PickZip();
        makeButton.Click += async (_,__) => await MakeVideoAsync();
    }

    private void PickPpt()
    {
        using var d = new OpenFileDialog { Filter = "PowerPoint (*.pptx;*.ppt)|*.pptx;*.ppt" };
        if (d.ShowDialog() == DialogResult.OK) { pptPath = d.FileName; pptBox.Text = pptPath; ValidateInputs(); }
    }

    private void PickZip()
    {
        using var d = new OpenFileDialog { Filter = "ZIP (*.zip)|*.zip" };
        if (d.ShowDialog() == DialogResult.OK) { zipPath = d.FileName; zipBox.Text = zipPath; ValidateInputs(); }
    }

    private void ValidateInputs()
    {
        makeButton.Enabled = pptPath != null && zipPath != null;
        status.Text = makeButton.Enabled ? "Đã chọn đủ dữ liệu. Bấm TẠO VIDEO MP4 để kiểm tra và xử lý." : "Hãy chọn PowerPoint và ZIP MP3.";
    }

    private async Task MakeVideoAsync()
    {
        if (pptPath is null || zipPath is null) return;
        using var save = new SaveFileDialog { Filter = "MP4 Video (*.mp4)|*.mp4", FileName = Path.GetFileNameWithoutExtension(pptPath) + ".mp4" };
        if (save.ShowDialog() != DialogResult.OK) return;

        makeButton.Enabled = false; progress.Value = 2;
        string work = Path.Combine(Path.GetTempPath(), "NTS_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(work);
        try
        {
            string audioDir = Path.Combine(work, "audio");
            string slideDir = Path.Combine(work, "slides");
            Directory.CreateDirectory(audioDir); Directory.CreateDirectory(slideDir);
            ZipFile.ExtractToDirectory(zipPath, audioDir);

            status.Text = "Đang đọc gói MP3...";
            var mp3s = Directory.GetFiles(audioDir, "*.mp3", SearchOption.AllDirectories)
                .Select(p => (path:p, n:ExtractNumber(Path.GetFileNameWithoutExtension(p))))
                .Where(x => x.n > 0).OrderBy(x => x.n).ToList();
            if (mp3s.Count == 0) throw new Exception("Không tìm thấy file MP3 có số thứ tự, ví dụ sl1.mp3, sl2.mp3...");

            status.Text = "Đang xuất slide bằng Microsoft PowerPoint...";
            int slideCount = ExportSlidesWithPowerPoint(pptPath, slideDir);
            progress.Value = 18;

            if (slideCount != mp3s.Count)
                throw new Exception($"PowerPoint có {slideCount} slide nhưng ZIP có {mp3s.Count} MP3.");
            for (int i=1;i<=slideCount;i++)
                if (!mp3s.Any(x=>x.n==i)) throw new Exception($"Thiếu file MP3 cho slide {i}.");

            string ffmpeg = FindTool("ffmpeg.exe"), ffprobe = FindTool("ffprobe.exe");
            var segments = new List<string>();
            for (int i=1;i<=slideCount;i++)
            {
                status.Text = $"Đang tạo đoạn {i}/{slideCount}...";
                var img = FindSlideImage(slideDir, i) ?? throw new Exception($"Không tìm thấy ảnh Slide {i}.");
                var audio = mp3s.First(x=>x.n==i).path;
                string seg = Path.Combine(work, $"seg_{i:0000}.mp4");
                await Run(ffmpeg, $"-y -loop 1 -i \"{img}\" -i \"{audio}\" -c:v libx264 -tune stillimage -pix_fmt yuv420p -vf \"scale=1920:1080:force_original_aspect_ratio=decrease,pad=1920:1080:(ow-iw)/2:(oh-ih)/2\" -c:a aac -b:a 192k -shortest -movflags +faststart \"{seg}\"");
                segments.Add(seg);
                progress.Value = 18 + (int)(72.0*i/slideCount);
            }

            string list = Path.Combine(work, "concat.txt");
            File.WriteAllLines(list, segments.Select(s => $"file '{s.Replace("'","'\\''")}'"));
            status.Text = "Đang nối video...";
            await Run(ffmpeg, $"-y -f concat -safe 0 -i \"{list}\" -c copy \"{save.FileName}\"");
            progress.Value = 100; status.Text = "Hoàn tất!";
            MessageBox.Show("Đã tạo MP4 thành công.", "NTS – Tạo MP4 từ PPT", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            status.Text = "Có lỗi.";
            MessageBox.Show(ex.Message, "NTS – Tạo MP4 từ PPT", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            makeButton.Enabled = true;
            try { Directory.Delete(work, true); } catch { }
        }
    }

    private static int ExtractNumber(string name)
    {
        var m = Regex.Match(name, @"(\d+)$");
        return m.Success && int.TryParse(m.Groups[1].Value, out int n) ? n : -1;
    }

    private static string? FindSlideImage(string dir, int i)
    {
        var candidates = new[] { $"Slide{i}.PNG", $"Slide{i}.png", $"Slide{i}.JPG", $"Slide{i}.jpg" };
        return candidates.Select(x=>Path.Combine(dir,x)).FirstOrDefault(File.Exists);
    }

    private static int ExportSlidesWithPowerPoint(string ppt, string outDir)
    {
        Type? t = Type.GetTypeFromProgID("PowerPoint.Application");
        if (t == null) throw new Exception("Máy chưa cài Microsoft PowerPoint.");
        dynamic? app = null, pres = null;
        try
        {
            app = Activator.CreateInstance(t)!;
            app.Visible = 0;
            pres = app.Presentations.Open(ppt, WithWindow: 0);
            int count = pres.Slides.Count;
            pres.Export(outDir, "PNG", 1920, 1080);
            pres.Close(); app.Quit();
            return count;
        }
        finally
        {
            try { if (pres != null) Marshal.FinalReleaseComObject(pres); } catch {}
            try { if (app != null) Marshal.FinalReleaseComObject(app); } catch {}
        }
    }

    private static string FindTool(string exe)
    {
        string local = Path.Combine(AppContext.BaseDirectory, exe);
        if (File.Exists(local)) return local;
        var paths = (Environment.GetEnvironmentVariable("PATH") ?? "").Split(';');
        var found = paths.Select(p=>Path.Combine(p.Trim(), exe)).FirstOrDefault(File.Exists);
        if (found != null) return found;
        throw new Exception($"Không tìm thấy {exe}. Bản phát hành GitHub phải kèm FFmpeg.");
    }

    private static async Task Run(string exe, string args)
    {
        var psi = new ProcessStartInfo(exe, args) { UseShellExecute=false, CreateNoWindow=true, RedirectStandardError=true, RedirectStandardOutput=true };
        using var p = Process.Start(psi) ?? throw new Exception("Không thể chạy công cụ tạo video.");
        string err = await p.StandardError.ReadToEndAsync();
        await p.WaitForExitAsync();
        if (p.ExitCode != 0) throw new Exception("Lỗi khi tạo video:\n" + err[^Math.Min(err.Length, 1800)..]);
    }
}
