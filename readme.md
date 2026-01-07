# 📦 Pokémon Mail Tools -- README

### Đọc Gmail (抽選結果・出荷メール) & iCloud bằng Node.js

### ✓ Dùng `pokemon_cre.json` làm Gmail Token

### ✓ Dùng App-Specific Password làm iCloud Token

------------------------------------------------------------------------

## 📁 Cấu trúc thư mục

    project-folder/
    │
    ├── check_ship_status.js      # Script check mail 出荷されました
    ├── gmail_check.js            # Script check mail Gmail (当選 / 抽選結果)
    ├── icloud_check.js           # Script check mail iCloud IMAP
    │
    ├── pokemon_cre.json          # Gmail OAuth Credentials (token dùng để xác thực)
    │
    ├── package.json
    └── readme.md

------------------------------------------------------------------------

# 1. 🔐 Gmail -- Lấy `pokemon_cre.json` (Gmail Token)

Google không cho dùng user/pass để đọc Gmail.\
Phải dùng OAuth theo chuẩn mới → file `pokemon_cre.json` chính là
**token + client secret**.

Dưới đây là hướng dẫn để tạo file đó.

------------------------------------------------------------------------

## 1.1. Tạo Project & Bật Gmail API

1.  Truy cập: https://console.cloud.google.com\
2.  Đăng nhập bằng Gmail bạn muốn đọc mail\
3.  Chọn **Select Project → New Project**\
4.  Đặt tên (ví dụ `Pokemon Gmail Tool`) → Create\
5.  Vào **APIs & Services → Library**\
6.  Tìm: **Gmail API**\
7.  Bấm **Enable**

------------------------------------------------------------------------

## 1.2. Thiết lập OAuth Consent Screen

1.  Vào **APIs & Services → OAuth consent screen**\
2.  User type → **External**\
3.  App name: tuỳ bạn\
4.  Thêm email bạn đang dùng vào phần **Test users**\
5.  Save & Publish

------------------------------------------------------------------------

## 1.3. Tạo OAuth Client (Desktop) → Tạo `pokemon_cre.json`

1.  Vào **APIs & Services → Credentials**\
2.  Bấm **Create Credentials → OAuth Client ID**\
3.  Application type → **Desktop App**\
4.  Nhấn **Create**\
5.  Nhấn **Download JSON**\
6.  Đổi tên thành:

```{=html}
<!-- -->
```
    pokemon_cre.json

✔ Đây là file token Gmail để chạy script\
✔ Không cần token.json nữa

Đặt file này vào cùng thư mục với:

-   `gmail_check.js`
-   `check_ship_status.js`

------------------------------------------------------------------------

# 2. 🍏 iCloud -- Lấy IMAP Token (App-Specific Password)

iCloud cho đọc mail qua IMAP nhưng **không dùng mật khẩu Apple ID**\
→ bắt buộc dùng **App-Specific Password**.

------------------------------------------------------------------------

## 2.1. Tạo iCloud IMAP Token

1.  Vào: https://appleid.apple.com\
2.  Đăng nhập\
3.  Đảm bảo **Two-Factor Authentication** đã bật\
4.  Vào **App-Specific Passwords**\
5.  Bấm **Generate Password**\
6.  Đặt tên (ví dụ: `node-imap`)\
7.  Apple trả về token dạng:

```{=html}
<!-- -->
```
    abcd-efgh-ijkl-mnop

✔ Đây chính là **iCloud IMAP Token**

Dùng token này trong file `icloud_check.js` để đăng nhập.

------------------------------------------------------------------------

# 3. ▶️ Run Project

## 3.1. Cài dependencies

``` bash
npm install
```

------------------------------------------------------------------------

## 3.2. Chạy Gmail Checker (抽選結果 -- 当選)

``` bash
node gmail_check.js
```

### Output:

-   In console danh sách mail
-   Xuất file CSV:

```{=html}
<!-- -->
```
    gmail_lottery_result.csv

------------------------------------------------------------------------

## 3.3. Chạy Check Ship (出荷されました)

``` bash
node check_ship_status.js
```

### Output:

-   In console 送り状番号 (WaybillNo)
-   Xuất file:

```{=html}
<!-- -->
```
    gmail_pokemon_shipping.csv

Format:

    Email;WaybillNo;TrackingUrl

------------------------------------------------------------------------

## 3.4. Chạy iCloud Checker

``` bash
node icloud_check.js
```

Dùng IMAP token (App-Specific Password) đã tạo ở bước 2.

------------------------------------------------------------------------

# 4. 📊 Công thức Excel So kết quả

### Excel tiếng Việt / Nhật (dùng dấu `;`)

    =IFERROR(VLOOKUP(A2; gmail_lottery_result!$A$2:$B$1000; 2; FALSE); "")
    =IFERROR(VLOOKUP(A2; icloud_lottery_result!$A$2:$B$1000; 2; FALSE); "")

### Excel tiếng Anh (dùng dấu `,`)

    =IFERROR(VLOOKUP(A2, gmail_lottery_result!$A$2:$B$1000, 2, FALSE), "")

------------------------------------------------------------------------

# 5. ⚠️ Lưu ý bảo mật

-   Không chia sẻ `pokemon_cre.json` cho người khác\
-   Nếu Gmail revoke quyền → thay file `pokemon_cre.json` mới\
-   Nếu iCloud đổi mật khẩu → tạo App-Specific Password mới

------------------------------------------------------------------------

# 6. 📌 Ghi chú thêm

-   Nếu bạn muốn gom tất cả script vào 1 CLI duy nhất → mình viết cho
    bạn\
-   Nếu cần bản **PDF hướng dẫn** → mình xuất PDF cho bạn\
-   Nếu muốn auto-update token Gmail → mình viết giúp luôn
