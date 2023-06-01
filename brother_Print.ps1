param (
    [Parameter(Mandatory = $false)]
    [int]$sod,
    [Parameter(Mandatory = $false)]
    [int]$tod,
    [Parameter(Mandatory = $true)]
    [string]$p
)

[void][System.Reflection.Assembly]::LoadFile($PSScriptRoot + '\Interop.bpac.dll')
Add-Type -Path "$PSScriptRoot\Interop.bpac.dll"
$Printers = New-Object bpac.PrinterClass


if ($sod) {
    $url = "http://192.168.25.235:8080/v1/slip/so/delivery/$sod"
}
elseif ($tod) {
    $url = "http://192.168.25.235:8080/v1/slip/to/delivery/$tod"
}
else {
    Write-Error "Please provide either -sod or -tod parameter."
    return
}

try {
    $Response = Invoke-WebRequest -Uri $url
    $StatusCode = $Response.StatusCode
}
catch {
    $StatusCode = $_.Exception.Response.StatusCode.value__
}

if ($StatusCode -eq 200) {
    $data = $Response.Content | ConvertFrom-Json

    $C_zip_code = "{0:###-####}" -f [double]$data.customer_zip_code
    $Customer_zip_code = "〒 "+$C_zip_code

    $Delivery_branch_code = $data.delivery_branch_code
    $Delivery_branch_code_1 = $Delivery_branch_code.Substring(0, 4)
    $Delivery_branch_code_2 = $Delivery_branch_code.Substring($Delivery_branch_code.Length - 3)
    $Customer_address = $data.customer_address_line

    $Customer_phone_number = $data.customer_phone_number
    if ($Customer_phone_number.Length -eq 10) {
        $Customer_phone_number_formatted = "0{0:##-####-####}" -f [double]$Customer_phone_number
    }
    elseif ($Customer_phone_number.Length -eq 11) {
        $Customer_phone_number_formatted = "{0:###-####-####}" -f [double]$Customer_phone_number
    }
    else {
        $Customer_phone_number_formatted = $Customer_phone_number
    }

    $Customer_name = $data.customer_address_line_3
    $Identifier = $data.identifier
    $Slip_number = $data.slip_number
    $Slip_number_text = "A$Slip_number" + "A"
    $Shipping_box_amount = $data.shipping_box_amount
    $Delivery_company_name = $data.delivery_company_name
    $Delivery_date = $data.delivery_date
    $Arrival_date = $data.arrival_date
    $Arrival_time = $data.arrival_time

    $Saloon_name = $data.saloon_name
    $S_zip_code = "{0:###-####}" -f [double]$data.saloon_zip_code
    $Saloon_zip_code = "〒 "+ $S_zip_code
    $Saloon_address_line_1 = $data.saloon_address_line_1
    $Saloon_address_line_2 = $data.saloon_address_line_2
    $Saloon_address = $Saloon_address_line_1 + $Saloon_address_line_2

    $Slip_msg = $data.slip_msg
    $S_phone_number = $data.saloon_phone_number
    if ($S_phone_number.Length -eq 10) {
        $Saloon_phone_number = "TEL: 0{0:##-####-####}" -f [double]$S_phone_number
    }
    elseif ($S_phone_number.Length -eq 11) {
        $Saloon_phone_number = "TEL: {###-####-####}" -f [double]$S_phone_number
    }
    else {
        $Saloon_phone_number = "TEL: " + $S_phone_number
    }

    $Printer_Name = $p

    $Label = New-Object bpac.DocumentClass
    $Filename = $PSScriptRoot + '\Template_v3.lbx'

    $foundPrinter = $Printers.GetInstalledPrinters() | Where-Object { $_ -like $Printer_Name }
    if ($foundPrinter) {
        if ($Label.Open($Filename)) {
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
            $Label.GetObject('saloon_phone_number').Text = $Saloon_phone_number

            # $Label.GetObject('delivery_msg').Text = $Delivery_msg
            $Label.GetObject('slip_msg').Text = $Slip_msg


            try {
                $Label.SetPrinter($foundPrinter, 0)
                $Label.StartPrint('', 0)
                $Label.PrintOut(1, 0)
                $Label.Close()
                $Label.EndPrint()
            }
            catch {
                Write-Output 'Failed'
                Write-Output $Label.ErrorCode
            }
        }
        else {
            Write-Output 'Failed to open label file'
        }
    }
    else {
        Write-Output 'Failed to find the specified printer'
    }


} else {
    Write-Error "Error: $StatusCode"
    return
}