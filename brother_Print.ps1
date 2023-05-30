[void][System.Reflection.Assembly]::LoadFile($PSScriptRoot+'\Interop.bpac.dll')
Add-Type -Path "$PSScriptRoot\Interop.bpac.dll"
$Printers = New-Object bpac.PrinterClass

$C_zip_code = $args[0]
$Customer_zip_code = "{0:###-####}" -f [double]$C_zip_code

$Delivery_branch_code = $args[1]
$Delivery_branch_code_1 = $Delivery_branch_code.Substring(0, 4)
$Delivery_branch_code_2 = $Delivery_branch_code.Substring($Delivery_branch_code.Length - 3)
$Customer_address = $args[2]
$Customer_phone_number = $args[3]

if ($Customer_phone_number.Length -eq 10) {
    $Customer_phone_number_formatted = "{0:###-####-####}" -f [double]$Customer_phone_number
} else {
    $Customer_phone_number_formatted = $Customer_phone_number
}


$Customer_name = $args[4]
$Identifier = $args[5]
$Slip_number = $args[6]
$Slip_number_text = "A$Slip_number" + "A"
$Shipping_box_amount = $args[7]
$Delivery_company_name = $args[8]
$Delivery_date = $args[9]
$Arrival_date = $args[10]
$Arrival_time = $args[11]

$Saloon_name = $args[12]
$S_zip_code = $args[13]
$Saloon_zip_code = "{0:###-####}" -f [double]$S_zip_code
$Saloon_address = $args[14]

# $Delivery_msg = $args[15]
$Slip_msg_unformatted = $args[15]

$Printer_Name = $args[16]

if (![string]::IsNullOrEmpty($Slip_msg_unformatted)) {
    $Slip_msg_c = [double]$Slip_msg_unformatted
    $Slip_msg = "¥{0:N0}" -f $Slip_msg_c
} else {
    $Slip_msg = ""
}


$Label = New-Object bpac.DocumentClass
$Filename = $PSScriptRoot+'\Template_v2.lbx'

$foundPrinter = $Printers.GetInstalledPrinters() | Where-Object { $_ -like $Printer_Name }
if ($foundPrinter) {
    if($Label.Open($Filename)) {
        $Label.GetObject('customer_zip_code').Text = $Customer_zip_code
        $Label.GetObject('delivery_branch_code').Text = $Delivery_branch_code
        $Label.GetObject('delivery_branch_code_1').Text = $Delivery_branch_code_1
        $Label.GetObject('delivery_branch_code_2').Text = $Delivery_branch_code_2
        $Label.GetObject('customer_address').Text = $Customer_address
        $Label.GetObject('customer_phone_number').Text = $Customer_phone_number_formatted
        $Label.GetObject('customer_name').Text = $Customer_name
        $Label.GetObject('identifier').Text = $Identifier
        $Label.GetObject('slip_number').Text = $Slip_number
        $Label.GetObject('slip_number_text').Text = $Slip_number_text
        $Label.GetObject('shipping_box_amount').Text = $Shipping_box_amount
        $Label.GetObject('delivery_company_name').Text = $Delivery_company_name
        $Label.GetObject('delivery_date').Text = $Delivery_date
        $Label.GetObject('arrival_date').Text = $Arrival_date
        $Label.GetObject('arrival_time').Text = $Arrival_time

        $Label.GetObject('saloon_name').Text = $Saloon_name
        $Label.GetObject('saloon_zip_code').Text = $Saloon_zip_code
        $Label.GetObject('saloon_address').Text = $Saloon_address

        # $Label.GetObject('delivery_msg').Text = $Delivery_msg
        $Label.GetObject('slip_msg').Text = $Slip_msg


        try {
            $Label.SetPrinter($foundPrinter, 0)
            $Label.StartPrint('',0)
            $Label.PrintOut(1, 0)
            $Label.Close()
            $Label.EndPrint()
        } catch {
            Write-Output 'Failed'
            Write-Output $Label.ErrorCode
        }
    } else {
        Write-Output 'Failed to open label file'
    }
} else {
    Write-Output 'Failed to find the specified printer'
}
