# MSsytem

## Introduction of user interface.

### Login
![1](https://user-images.githubusercontent.com/73460497/193453780-5b27d8bc-e28d-47be-9b6e-cddc10a3635f.jpg)

- 根據帳號級別分成管理者跟使用者，在前端設置簡單的密碼保護
- According to account levels, divide into administrators and users, and set up password protection on the front end.
### Main Page
![2](https://user-images.githubusercontent.com/73460497/193453866-b5de5be4-7392-44f2-a541-3477edc75e87.jpg)

- 介面以乾淨、俐落、簡單為主，考量到任何人能快速上手，主要分成新工單、啟動、停機、參數設定、結束工單
- The user interface is clean, tidy, and simple, ensuring that anyone can quickly get started. It is mainly divided into new work order, start, stop, parameter setting, and end work order sections.

#### 新工單：每一批為一個工單，必須在新的一批開始之前新增一個工單(如下圖)
#### New Work Order: Need to new a work order before starting a new process (as shown below).
![3](https://user-images.githubusercontent.com/73460497/193454279-651060f0-5491-42de-b828-82aae7908bb4.jpg)

- 點選後需輸入產品批次號、選擇作業員編號(可於參數設定任意增減、刪改)、早中晚班
- After clicking work order, you need to enter the product batch number, select the operator number (which can be adjusted or modified arbitrarily in parameter settings), and choose the shift (morning, afternoon, or night).

- 確定後便會在D槽建立相關LOG檔，如下圖
- After confirmation, relevant log files will be created in drive D, as shown below.

![5](https://user-images.githubusercontent.com/73460497/193454338-47793ccc-6322-4085-9da8-903952bfaa76.jpg)

![4](https://user-images.githubusercontent.com/73460497/193454519-355f1b2c-f677-49cc-b725-8477af2ce9d2.jpg)

- 並在右上角顯示目前與機台(PLC)連線狀態
- And display the current connection status with the machine (PLC) in the top right corner.

#### 啟動
#### Start

- 點選啟動會確認是否已開啟工單，並建立與PLC連線。
- Clicking on "啟動" will confirm whether a work order has been opened and establish a connection with the PLC.

#### 停機
#### Stop

- 緊急按鈕，一般都是從人機介面或機器的緊急按鈕停機，此功能作為備用
- The emergency button serves as a backup function, as typically, shutdowns are initiated from the human-machine interface or the machine's emergency button.

#### 參數設定
#### Parameter Settings
![7](https://user-images.githubusercontent.com/73460497/193454549-d6f28031-2e1d-472a-98e1-dbca729cc567.jpg)

- 可新增刪除作業員編號、改變規格上下限，管理員帳號密碼、使用者密碼、故障排除、急停的原因
- You can add or delete operator numbers, modify specification upper and lower limits, manage administrator account password, user password, troubleshooting, and emergency stop reasons.

![8](https://user-images.githubusercontent.com/73460497/193454601-2b1c3d7d-0fc9-44ca-8758-133b8a6cd400.jpg)

- 也可新增刪除停機原因
- You can also add or delete shutdown reasons.

- 若停機必須由管理員來處理
- Shutdowns must be handled by administrators.

![9](https://user-images.githubusercontent.com/73460497/193454632-8e4a19a9-ad5e-4e3c-b36b-402594f31035.jpg)

#### 結束工單
#### End Work Order

- 於每批結束後點選，便會儲存所有資料，若在機器運作被斷電，資料都有暫存在database裡面，重新啟動會顯示需先儲存在開始新的或繼續舊工單
- After each work order ends, clicking on it will save all data. If there's a power outage during machine operation, the data will be temporarily stored in the database. Upon restarting, a prompt will appear to either save as a new batch or continue with the previous work order.

![10](https://user-images.githubusercontent.com/73460497/193454771-0a5fb5f1-5fe0-4597-ae3f-7577fd0e5699.jpg)

- 並設有防呆提醒，避免使用者尚未儲存而造成資料流失
- Also, there is a preventive reminder to avoid data loss caused by users not saving their work order.
