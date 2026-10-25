# BK-Textile DSS: Tối ưu hóa kế hoạch sản xuất và quản lý phân phối đa chi nhánh

Hệ thống hỗ trợ ra quyết định (DSS) tích hợp **Excel nâng cao, VBA UserForm và Python (Quy hoạch tuyến tính)** cho công ty may mặc BK-Textile. Dự án thực hiện theo đề thí nghiệm *"Tối ưu hóa kế hoạch sản xuất và quản lý phân phối đa chi nhánh"*.

## 1. Bối cảnh bài toán

BK-Textile có **03 nhà máy** (F1 TP.HCM, F2 Đồng Nai, F3 Long An) sản xuất **03 dòng sản phẩm** (Áo thun AT001, Áo sơ mi AS002, Quần Tây QT003) và cung cấp cho **02 trung tâm phân phối** (DC) tại miền Bắc (North) và miền Nam (South). Việc điều phối sản xuất và vận chuyển hiện làm thủ công, dẫn đến chi phí cao và sai sót dữ liệu.

DSS cần đáp ứng:

- Có các bảng dữ liệu gốc được chuẩn hóa cấu trúc.
- Có giao diện (UserForm) để nhập, chỉnh sửa, xóa dữ liệu; **không** cho người dùng nhập trực tiếp lên bảng dữ liệu gốc.
- Tự động tính toán để tạo thông báo, cảnh báo, báo cáo; người dùng chỉ cần nhập liệu qua giao diện hoặc cung cấp file dữ liệu đầu vào.

## 2. Mục tiêu

- **Chuyên môn:** xây dựng mô hình Quy hoạch tuyến tính (LP) để giảm thiểu tổng chi phí sản xuất và vận chuyển.
- **Kỹ thuật:** sử dụng thành thạo các công cụ xử lý dữ liệu nâng cao trên Excel, lập trình giao diện UserForm bằng VBA và giải toán tối ưu bằng Python.

## 3. Dữ liệu đầu vào

| File | Nội dung |
|---|---|
| `Master_Data.xlsx` | Sheet `Product_BOM`: SKU, định mức vải (m/sp), công lao động (giờ/sp), giá bán. Sheet `Factory_Capacity`: công suất tối đa (giờ), chi phí vận hành ($/giờ) |
| `Inventory_System.xlsx` | Sheet `Inventory_Log`: nhật ký nhập/xuất kho. Sheet `Stock_Level`: tồn kho hiện tại và mức an toàn |
| `Order_History.xlsx` | 500 dòng lịch sử đơn hàng (tháng 1–3/2025): mã đơn, ngày, khách hàng, SKU, số lượng, khu vực |
| `Demand_Forecast.csv` | Nhu cầu dự báo theo SKU và khu vực |
| `Transport_Costs.csv` | Chi phí vận chuyển từ mỗi nhà máy đến mỗi DC |

**Lưu ý chất lượng dữ liệu của `Order_History.xlsx`:** có 20 dòng trùng hoàn toàn và khoảng 29 dòng mã SKU viết chữ thường (`at001`, `as002`, `qt003`). Cần chuẩn hóa (`UPPER`, loại dòng trùng) trước khi dùng để dự báo hoặc làm Pivot Table.

## 4. Cấu trúc thư mục

```
.
├── README.md
├── optimize.py                  
├── data/                        
│   ├── Master_Data.xlsx
│   ├── Inventory_System.xlsx
│   ├── Demand_Forecast.csv
│   └── Transport_Costs.csv
├── Solution.xlsx          
└── VBA
```

## 5. Nhiệm vụ chi tiết

### Nhiệm vụ 1: Thiết kế hệ thống quản lý dữ liệu (Excel nâng cao)

**Cấu trúc dữ liệu.** Xây dựng các bảng chuẩn hóa: Danh mục sản phẩm, Định mức nguyên vật liệu (BOM), Công suất nhà máy, Chi phí vận chuyển giữa các nút.

**Xử lý dữ liệu.**

- Dùng **Data Validation** (List, Custom) để kiểm soát dữ liệu đầu vào.
- Dùng **hàm cơ sở dữ liệu** (`DSUM`, `DAVERAGE`) và **Advanced Filter** để trích xuất thông tin đơn hàng theo điều kiện.
- Dùng công cụ **Consolidate** để hợp nhất dữ liệu doanh số từ các chi nhánh.
- Dùng **Pivot Table và Slicer** để xây dựng Dashboard báo cáo tình hình sản xuất theo tháng/quý.

Trạng thái:

- [ ] Bảng dữ liệu chuẩn hóa (Product, BOM, Factory, Transport)
- [ ] Data Validation
- [ ] DSUM / DAVERAGE và Advanced Filter
- [ ] Consolidate
- [ ] Pivot Table, Slicer và Dashboard

### Nhiệm vụ 2: Xây dựng phần mềm mini quản trị kho (VBA-Excel)

**Thiết kế UserForm.** Giao diện nhập liệu gồm `TextBox` (nhập số lượng), `ComboBox` (chọn nhà máy/sản phẩm) và các `CommandButton`.

**Chức năng.**

- Tự động tìm dòng trống cuối cùng để lưu dữ liệu từ Form vào Database.
- Dùng cấu trúc `If...Then` và `MsgBox` để cảnh báo khi tồn kho chạm mức an toàn (lấy từ `Stock_Level`).
- Dùng vòng lặp `For...Next` để tự động định dạng và kẻ bảng báo cáo sau khi xuất dữ liệu.

Trạng thái:

- [ ] UserForm nhập/sửa/xóa dữ liệu
- [ ] Lưu vào dòng trống cuối cùng của `Inventory_Log`
- [ ] Cảnh báo tồn kho an toàn
- [ ] Tự động định dạng và kẻ bảng báo cáo

### Nhiệm vụ 3: Tối ưu hóa kế hoạch (Quy hoạch tuyến tính và Python)

Trạng thái: **đã hoàn thành** (`optimize.py`).

#### 5.1. Mô hình toán học

**Tập hợp và chỉ số**

- $i \in I = \{F1, F2, F3\}$: nhà máy
- $j \in J = \{North, South\}$: trung tâm phân phối
- $k \in K = \{AT001, AS002, QT003\}$: sản phẩm

**Tham số**

| Ký hiệu | Ý nghĩa | Nguồn |
|---|---|---|
| $D_{jk}$ | Nhu cầu dự báo của sản phẩm $k$ tại DC $j$ | `Demand_Forecast.csv` |
| $Q_k$ | Tồn kho hiện tại của sản phẩm $k$ | `Stock_Level` |
| $S_k$ | Mức tồn kho an toàn của sản phẩm $k$ | `Stock_Level` |
| $Cap_i$ | Công suất tối đa (giờ) của nhà máy $i$ | `Factory_Capacity` |
| $C_i^{prod}$ | Chi phí vận hành ($/giờ) của nhà máy $i$ | `Factory_Capacity` |
| $T_k$ | Công lao động (giờ/sp) | `Product_BOM` |
| $M_k$ | Định mức vải (m/sp), dùng để báo cáo lượng vải sử dụng | `Product_BOM` |
| $C_{ij}^{trans}$ | Chi phí vận chuyển ($/sp) từ nhà máy $i$ đến DC $j$ | `Transport_Costs.csv` |

**Biến quyết định**

- $X_{ijk} \ge 0$: số lượng sản phẩm $k$ sản xuất tại nhà máy $i$ và phân phối cho DC $j$.
- $\rho \in [0, 1]$: tỷ lệ sử dụng công suất chung của các nhà máy (giờ công đã dùng / công suất tối đa). $\rho$ không được nhập vào mà được xác định trong quá trình giải mô hình.

**Hàm mục tiêu:** tối thiểu hóa tổng chi phí sản xuất và vận chuyển.

$$\min Z = \sum_{i \in I}\sum_{j \in J}\sum_{k \in K} X_{ijk}\left(T_k \cdot C_i^{prod} + C_{ij}^{trans}\right)$$

**Ràng buộc**

1. Đảm bảo mức tồn kho an toàn:

$$Q_k + \sum_{i \in I}\sum_{j \in J} X_{ijk} - \sum_{j \in J} D_{jk} \ge S_k, \quad \forall k \in K$$

2. Cân bằng tải và giới hạn công suất giữa các nhà máy (tỷ lệ sử dụng công suất của mọi nhà máy bằng nhau và không vượt quá 1):

$$\sum_{j \in J}\sum_{k \in K} T_k \cdot X_{ijk} = \rho \cdot Cap_i, \quad \forall i \in I$$

3. Đáp ứng nhu cầu theo từng khu vực:

$$\sum_{i \in I} X_{ijk} \ge D_{jk}, \quad \forall j \in J,\ k \in K$$

#### 5.2. Giả định và hạn chế

- Toàn bộ chi phí trong mô hình được quy ước cùng đơn vị tiền tệ là **$**, nên hàm mục tiêu cộng trực tiếp chi phí sản xuất và vận chuyển. Do cước vận chuyển có độ lớn hàng nghìn so với chi phí sản xuất vài chục đến vài nghìn, kết quả chủ yếu do chi phí vận chuyển quyết định.
- Bộ dữ liệu **không cung cấp chi phí lưu kho** nên hàm mục tiêu chỉ gồm chi phí sản xuất và vận chuyển.
- Bộ dữ liệu **không cung cấp lượng vải sẵn có**, nên định mức vải $M_k$ chỉ được dùng để tính tổng lượng vải cần dùng ở kết quả, không đưa vào ràng buộc.
- Tồn kho hiện tại và tồn kho an toàn chỉ có theo sản phẩm (không theo vùng). Vì vậy nhu cầu từng vùng được đáp ứng hoàn toàn bằng sản xuất mới (ràng buộc 3), còn tồn kho dùng cho mức an toàn (ràng buộc 1).
- Chi phí vận chuyển được giả định như nhau cho mọi SKU trên cùng một tuyến.
- Kỳ kế hoạch chưa được nêu trong dữ liệu. Công suất và nhu cầu được hiểu là cùng một kỳ.

#### 5.3. Cách chạy

Yêu cầu Python 3.9 trở lên.

```bash
pip install pandas numpy scipy openpyxl
python optimize.py data Ket_qua_toi_uu.xlsx
```

Tham số dòng lệnh: thư mục dữ liệu (mặc định `data`) và tên file kết quả (mặc định `Ket_qua_toi_uu.xlsx`). Thư mục dữ liệu cần có đủ 4 file `Master_Data.xlsx`, `Inventory_System.xlsx`, `Demand_Forecast.csv`, `Transport_Costs.csv`.

Chương trình dùng `dict`/`list` để lưu tham số, đọc và chuẩn hóa dữ liệu đầu vào, giải LP bằng `scipy.optimize.linprog` (HiGHS) và xuất kết quả ra Excel.

#### 5.4. Kết quả đầu ra

File Excel gồm 6 sheet: `Tổng quan`, `Kế hoạch sản xuất`, `Lộ trình vận chuyển`, `Sử dụng nhà máy`, `Kiểm tra nhu cầu`, `Kiểm tra tồn kho`.

Kết quả với bộ dữ liệu hiện tại:

| Chỉ tiêu | Giá trị |
|---|---|
| Tổng chi phí tối ưu $Z$ | ≈ 66,11 triệu $ |
| Chi phí sản xuất | ≈ 62,6 nghìn $ |
| Chi phí vận chuyển | ≈ 66,05 triệu $ |
| Tỷ lệ sử dụng công suất $\rho$ | 54,9% |
| Giờ công sử dụng F1 / F2 / F3 | 2.744 / 2.470 / 2.196 |
| Tổng sản lượng | 7.400 sp (bằng tổng nhu cầu) |
| Tổng vải cần dùng | 12.080 m |

Do không có chi phí lưu kho nên mô hình sản xuất đúng bằng nhu cầu. Tồn kho đầu kỳ (100 / 50 / 30) lớn hơn mức an toàn (50 / 30 / 20) nên ràng buộc 1 chưa chặt với bộ dữ liệu này.

## 6. Công nghệ sử dụng

- Microsoft Excel (Data Validation, hàm cơ sở dữ liệu, Advanced Filter, Consolidate, Pivot Table, Slicer)
- VBA (UserForm, `If...Then`, `MsgBox`, `For...Next`)
- Python: `pandas`, `numpy`, `scipy`, `openpyxl`
