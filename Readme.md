Brother Print Script - Readme
Made by : Ali Ahammed Rohid 
For : Kikuya Bisyodo Inc.
==============================================


This script utilizes the Brother BPAC library to print labels using a Brother printer. Please provide the following arguments when executing the script:

1. Customer Zip Code: The zip code of the customer. Format it as a 7-digit number without hyphens. For example: 1112222.

2. Delivery Branch Code: Barcode for the delivery branch code. It should be an alphanumeric code. For example: 1111000.

3. Customer Address: The address of the customer. For example: "東京都渋谷区恵比寿西１－１３－７ ＫＢ東京都渋谷区東京都渋谷区恵比寿西１－１３－７ ＫＢ東京都渋谷区東京都渋谷区恵比寿西１－１３－７ ＫＢ東京都渋谷区".

4. Customer Phone Number: The phone number of the customer. Format it as a 10-digit number without hyphens. For example: 08011112222.

5. Customer Name: The name of the customer. For example: "John Doe".

6. Identifier: An identifier. For example: "佐川".

7. Slip Number: Center Barcode. For example: "121212121212".

8. Shipping Box Amount: The number of shipping boxes. For example: "5".

9. Delivery Company Name: The name of the delivery company. For example: "Delivery Company(LC)".

10. Delivery Date: The delivery date. Format it as a date in the format YYYY/MM/DD. For example: "2023/06/01".

11. Arrival Date: The arrival date. Format it as a date in the format YYYY/MM/DD. For example: "2023/06/02".

12. Arrival Time: The arrival time. For example: "12時～14時".

13. Saloon Name: The name of the saloon. For example: "Saloon Name".

14. Saloon Zip Code: The zip code of the saloon. Format it as a 7-digit number without hyphens. For example: 3334444.

15. Saloon Address: The address of the saloon. For example: "東京都渋谷区恵比寿西東京都渋谷区恵比寿西１－１３－７".

16. Slip Message: An optional slip message. If provided, it will be formatted as Japanese currency and preceded by the yen sign. For example: "49060".

17. Printer Name: The name of the printer. For example: "Brother TD-4420DN".

Example Commands : 

.\brother_Print.ps1 "1112222" "1111000" "東京都渋谷区恵比寿西１－１３－７ＫＢ東京都渋谷区東京都渋谷区恵比寿西１－１３－７ＫＢ東京都渋谷区東京都渋谷区恵比寿西１－１３－７ＫＢ東京都渋谷区" "08011112222" "John Doe" "佐川" "121212121212" "5" "佐川代引 (LC)" "2023/06/01" "2023/06/02" "12時～14時" "株式会社きくや美粧堂サロンあいうえおかき" "3334444" "東京都渋谷区恵比寿西東京都渋谷区恵比寿西１－１３－７" "49060" "Brother TD-4420DN"

powershell.exe -ExecutionPolicy Bypass -File fileName.ps1 arg[0] arg[1] ... arg[n]