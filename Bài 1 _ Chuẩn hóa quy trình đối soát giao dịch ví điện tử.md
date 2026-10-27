## PHẦN 1. PHÂN TÍCH

Trong bảng Excel giao dịch của ví điện tử PayKX:

-   Record (bản ghi): Là một dòng dữ liệu, thể hiện thông tin của một giao dịch.
    
-   Field (trường dữ liệu): Là một cột thông tin, ví dụ: Mã giao dịch, Mã người dùng, Tên người dùng, Số tiền, Phí giao dịch và Ngày giao dịch.
    

 ## PHẦN 2. TRIỂN KHAI

1.  File Word – Tài liệu SRS
    

Em tạo tài liệu SRS mô tả chức năng hoàn tiền của ví điện tử PayKX, gồm 5 trang:

-   Trang 1: Bìa tài liệu và mục lục tự động (Auto TOC).
    
-   Trang 2: Giới thiệu dự án và mục tiêu chức năng hoàn tiền.
    
-   Trang 3: Yêu cầu chức năng hoàn tiền, sử dụng Heading 1 và Heading 2 để phân cấp nội dung.
    
-   Trang 4: Quy trình hoàn tiền, điều kiện thực hiện và xử lý lỗi.
    
-   Trang 5: Yêu cầu bảo mật, kiểm thử và kết luận.
    

Em sử dụng Heading Styles để định dạng tiêu đề và tạo mục lục tự động, giúp tài liệu rõ ràng, dễ theo dõi.

2.  File Excel – Bảng đối soát giao dịch
    

Em tạo bảng dữ liệu gồm các cột: Mã giao dịch, Mã người dùng, Tên người dùng, Số tiền và Phí giao dịch.

-   Sử dụng hàm TRIM để loại bỏ khoảng trắng thừa trong mã giao dịch.
    
-   Sử dụng hàm VLOOKUP với tham số FALSE để tìm chính xác tên người dùng theo mã giao dịch.
    
-   Không gộp ô (Merge Cell), không để hàng trống giữa các bản ghi.
    
-   Mỗi cột chỉ chứa một loại dữ liệu để thuận tiện cho việc lọc và đối soát.
    

Ví dụ công thức tra cứu tên người dùng:

=VLOOKUP(TRIM(A2),$F$2:$G$100,2,FALSE)

Trong đó, A2 là mã giao dịch cần tra cứu; cột F chứa mã giao dịch trong bảng tra cứu và cột G chứa tên người dùng.

3.  File PowerPoint – Báo cáo doanh thu
    

Em thiết kế bài trình bày báo cáo lỗ lãi của PayKX theo nguyên tắc 10/20/30 của Guy Kawasaki:

-   Tối đa 10 slide.
    
-   Thời gian trình bày khoảng 20 phút.
    
-   Cỡ chữ tối thiểu 30pt.
    

Em sử dụng tiêu đề khẳng định (Assertion Title), ví dụ: “Doanh thu tăng 15% nhờ ưu đãi nạp điện thoại”.

Nội dung gồm doanh thu, chi phí, phí ngân hàng, lợi nhuận và đề xuất cải thiện. Em sử dụng biểu đồ cột hoặc biểu đồ đường để thể hiện sự biến động dòng tiền, giúp người xem dễ hiểu và so sánh số liệu.

 ## PHẦN 3. KIỂM THỬ

Khi hàm VLOOKUP trả về lỗi #N/A, em thực hiện kiểm tra theo 4 bước:

Bước 1: Kiểm tra mã giao dịch cần tra cứu có tồn tại trong bảng dữ liệu hay không.

Bước 2: Sử dụng hàm TRIM để loại bỏ khoảng trắng thừa trong mã giao dịch.

Bước 3: Kiểm tra vùng tra cứu và bảo đảm cột chứa mã giao dịch nằm ở vị trí đầu tiên của vùng dữ liệu.

Bước 4: Kiểm tra tham số FALSE để bảo đảm tra cứu chính xác, sau đó chạy lại công thức và kiểm tra kết quả.

## Em kết luận như sau :

Qua bài tập, em biết cách tạo tài liệu SRS bằng Word, đối soát giao dịch bằng Excel và trình bày báo cáo doanh thu bằng PowerPoint. Các công cụ này giúp em quản lý dữ liệu chính xác, trình bày thông tin rõ ràng và phát hiện lỗi trong quá trình đối soát.
