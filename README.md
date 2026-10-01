# NTS – Tạo MP4 từ PPT

**Tác giả: Nguyễn Thanh Sang**

V1 dành cho Windows:
- Chọn PPT/PPTX.
- Chọn ZIP chứa MP3 theo thứ tự `sl1.mp3`, `sl2.mp3`, ...
- Tự kiểm tra số slide và số MP3.
- Microsoft PowerPoint xuất từng slide thành hình tĩnh.
- FFmpeg ghép từng slide với MP3 tương ứng và nối thành MP4 Full HD 1080p.
- Không sửa file PowerPoint gốc.
- V1 không phát Animation/Transition.

## Điều kiện trên máy sử dụng
- Windows 10/11 64-bit.
- Microsoft PowerPoint đã cài đặt.

## Build
GitHub Actions sẽ tự build khi project được đưa lên nhánh `main`.
Sau khi workflow `Build Windows EXE` hoàn tất, tải Artifact `NTS-Tao-MP4-V1-Windows`.
