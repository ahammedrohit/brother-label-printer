param (
    [Parameter(Mandatory = $true)]
    [string]$u,
    [Parameter(Mandatory = $true)]
    [string]$p
)

[void][System.Reflection.Assembly]::LoadFile($PSScriptRoot + '\Interop.bpac.dll')
Add-Type -Path "$PSScriptRoot\Interop.bpac.dll"
$Printers = New-Object bpac.PrinterClass


if ($u) {
    $url = $u
    Write-Output $url 
}
else {
    Write-Error "Please provide the order ID parameter."
    return
}

try {
    $Response = Invoke-WebRequest -Uri $url
    $StatusCode = $Response.StatusCode
}
catch {
    $StatusCode = $_.Exception.Response.StatusCode.value__
}

# if ($StatusCode -eq 200) {
#     $data = $Response.Content | ConvertFrom-Json

#     $PrintJobs = @()

#     foreach ($item in $data) {
#         $Customer_zip_code = $item.customer_zip_code
#         $Delivery_branch_code = $item.delivery_branch_code
#         $Delivery_branch_code_1 = $item.delivery_branch_code_1
#         $Delivery_branch_code_2 = $item.delivery_branch_code_2

#         if ($item.customer_address_line_1.Length -gt 25) {
#             $Customer_address_line_l1 = $item.customer_address_line_1
#             $Customer_address_line_s1 = ""
#         } else {
#             $Customer_address_line_l1 = ""
#             $Customer_address_line_s1 = $item.customer_address_line_1
#         }
#         if ($item.customer_address_line_2.Length -gt 25) {
#             $Customer_address_line_l2 = $item.customer_address_line_2
#             $Customer_address_line_s2 = ""
#         } else {
#             $Customer_address_line_l2 = ""
#             $Customer_address_line_s2 = $item.customer_address_line_2
#         }
#         if ($item.customer_address_line_3.Length -gt 25) {
#             $Customer_address_line_l3 = $item.customer_address_line_3
#             $Customer_address_line_s3 = ""
#         } else {
#             $Customer_address_line_l3 = ""
#             $Customer_address_line_s3 = $item.customer_address_line_3
#         }
#         $Customer_phone_number = $item.customer_phone_number
#         $Identifier = $item.identifier
#         $Slip_number = $item.slip_number
#         if ($Slip_number) {
#             $Slip_number_text = "A" + "$Slip_number" + "A"
#         } else {
#             $Slip_number_text = ""
#         }

#         $Shipping_box_amount = $item.shipping_box_amount
#         $Delivery_company_name = $item.delivery_company_name
#         $Delivery_date = $item.delivery_date
#         $Arrival_date = $item.arrival_date
#         $Arrival_time = $item.arrival_time
#         $Saloon_name = $item.saloon_name
#         $S_zip_code = $item.saloon_zip_code
#         $Saloon_zip_code = $S_zip_code
#         $Saloon_address_line_1 = $item.saloon_address_line_1
#         $Saloon_address_line_2 = $item.saloon_address_line_2
#         $Slip_msg = $item.slip_msg
#         $Saloon_phone_number = "TEL: " + $item.saloon_phone_number

#         $Printer_Name = $p

#         $Label = New-Object bpac.DocumentClass
#         $Filename = $PSScriptRoot + '\Template_v3.lbx'
#         $LabelFilename = $PSScriptRoot + '\' + (Get-Date -Format "yyyyMMddHHmmss") + '.lbx'

#         $foundPrinter = $Printers.GetInstalledPrinters() | Where-Object { $_ -like $Printer_Name }
#         if ($foundPrinter) {
#             if ($Label.Open($Filename)) {
#                 $Label.GetObject('customer_zip_code').Text = $Customer_zip_code
#                 $Label.GetObject('delivery_branch_code').Text = $Delivery_branch_code
#                 $Label.GetObject('delivery_branch_code_1').Text = $Delivery_branch_code_1
#                 $Label.GetObject('delivery_branch_code_2').Text = $Delivery_branch_code_2
#                 $Label.GetObject('customer_address_line_l1').Text = $Customer_address_line_l1
#                 $Label.GetObject('customer_address_line_l2').Text = $Customer_address_line_l2
#                 $Label.GetObject('customer_address_line_l3').Text = $Customer_address_line_l3
#                 $Label.GetObject('customer_address_line_s1').Text = $Customer_address_line_s1
#                 $Label.GetObject('customer_address_line_s2').Text = $Customer_address_line_s2
#                 $Label.GetObject('customer_address_line_s3').Text = $Customer_address_line_s3
#                 $Label.GetObject('customer_phone_number').Text = $Customer_phone_number
#                 $Label.GetObject('identifier').Text = $Identifier
#                 $Label.GetObject('slip_number').Text = $Slip_number
#                 $Label.GetObject('slip_number_text').Text = $Slip_number_text
#                 $Label.GetObject('shipping_box_amount').Text = $Shipping_box_amount
#                 $Label.GetObject('delivery_company_name').Text = $Delivery_company_name
#                 $Label.GetObject('delivery_date').Text = $Delivery_date
#                 $Label.GetObject('arrival_date').Text = $Arrival_date
#                 $Label.GetObject('arrival_time').Text = $Arrival_time

#                 $Label.GetObject('saloon_name').Text = $Saloon_name
#                 $Label.GetObject('saloon_zip_code').Text = $Saloon_zip_code
#                 $Label.GetObject('saloon_address_line_1').Text = $Saloon_address_line_1
#                 $Label.GetObject('saloon_address_line_2').Text = $Saloon_address_line_2
#                 $Label.GetObject('saloon_phone_number').Text = $Saloon_phone_number

#                 $Label.GetObject('delivery_msg').Text = $Delivery_msg
#                 $Label.GetObject('slip_msg').Text = $Slip_msg

#                 try {
#                     $Label.SetPrinter($foundPrinter, 0)
#                     $Label.StartPrint($LabelFilename, 0)
#                     $PrintJobs += $Label.PrintOut(1, 0)
#                     $Label.Close()
#                     $Label.EndPrint()
#                 }
#                 catch {
#                     Write-Output 'Failed'
#                     Write-Output $Label.ErrorCode
#                 }
#             }
#             else {
#                 Write-Output 'Failed to open the label file'
#             }
#         }
#         else {
#             Write-Output 'Failed to find the specified printer'
#         }
#     }

#     if ($PrintJobs.Count -gt 0) {
#         Write-Output "Print Jobs: $($PrintJobs -join ', ')"
#     }
#     else {
#         Write-Output 'No print jobs created'
#     }
# }
# else {
#     Write-Error "Error: $StatusCode"
#     return
# }

if ($StatusCode -eq 200) {
    $data = $Response.Content | ConvertFrom-Json

    $Label = New-Object bpac.DocumentClass
    $Filename = $PSScriptRoot + '\Template_v3.lbx'
    $LabelFilename = $PSScriptRoot + '\' + (Get-Date -Format "yyyyMMddHHmmss") + '.lbx'

    $foundPrinter = $Printers.GetInstalledPrinters() | Where-Object { $_ -like $p }
    if ($foundPrinter) {
        if ($Label.Open($Filename)) {
            foreach ($item in $data) {
                $Customer_zip_code = $item.customer_zip_code
                $Delivery_branch_code = $item.delivery_branch_code
                $Delivery_branch_code_1 = $item.delivery_branch_code_1
                $Delivery_branch_code_2 = $item.delivery_branch_code_2

                if ($item.customer_address_line_1.Length -gt 25) {
                    $Customer_address_line_l1 = $item.customer_address_line_1
                    $Customer_address_line_s1 = ""
                }
                else {
                    $Customer_address_line_l1 = ""
                    $Customer_address_line_s1 = $item.customer_address_line_1
                }
                if ($item.customer_address_line_2.Length -gt 25) {
                    $Customer_address_line_l2 = $item.customer_address_line_2
                    $Customer_address_line_s2 = ""
                }
                else {
                    $Customer_address_line_l2 = ""
                    $Customer_address_line_s2 = $item.customer_address_line_2
                }
                if ($item.customer_address_line_3.Length -gt 25) {
                    $Customer_address_line_l3 = $item.customer_address_line_3
                    $Customer_address_line_s3 = ""
                }
                else {
                    $Customer_address_line_l3 = ""
                    $Customer_address_line_s3 = $item.customer_address_line_3
                }
                $Customer_phone_number = $item.customer_phone_number
                $Identifier = $item.identifier
                $Slip_number = $item.slip_number
                if ($Slip_number) {
                    $Slip_number_text = "A" + "$Slip_number" + "A"
                }
                else {
                    $Slip_number_text = ""
                }

                $Shipping_box_amount = $item.shipping_box_amount
                $Delivery_company_name = $item.delivery_company_name
                $Delivery_date = $item.delivery_date
                $Arrival_date = $item.arrival_date
                $Arrival_time = $item.arrival_time
                $Saloon_name = $item.saloon_name
                $S_zip_code = $item.saloon_zip_code
                $Saloon_zip_code = $S_zip_code
                $Saloon_address_line_1 = $item.saloon_address_line_1
                $Saloon_address_line_2 = $item.saloon_address_line_2
                $Slip_msg = $item.slip_msg
                $Saloon_phone_number = "TEL: " + $item.saloon_phone_number

                $Label.GetObject('customer_zip_code').Text = $Customer_zip_code
                $Label.GetObject('delivery_branch_code').Text = $Delivery_branch_code
                $Label.GetObject('delivery_branch_code_1').Text = $Delivery_branch_code_1
                $Label.GetObject('delivery_branch_code_2').Text = $Delivery_branch_code_2
                $Label.GetObject('customer_address_line_l1').Text = $Customer_address_line_l1
                $Label.GetObject('customer_address_line_l2').Text = $Customer_address_line_l2
                $Label.GetObject('customer_address_line_l3').Text = $Customer_address_line_l3
                $Label.GetObject('customer_address_line_s1').Text = $Customer_address_line_s1
                $Label.GetObject('customer_address_line_s2').Text = $Customer_address_line_s2
                $Label.GetObject('customer_address_line_s3').Text = $Customer_address_line_s3
                $Label.GetObject('customer_phone_number').Text = $Customer_phone_number
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
                $Label.GetObject('saloon_address_line_1').Text = $Saloon_address_line_1
                $Label.GetObject('saloon_address_line_2').Text = $Saloon_address_line_2
                $Label.GetObject('saloon_phone_number').Text = $Saloon_phone_number
                $Label.GetObject('delivery_msg').Text = $Delivery_msg
                $Label.GetObject('slip_msg').Text = $Slip_msg

                try {
                    $Label.SetPrinter($foundPrinter, 0)
                    $Label.StartPrint($LabelFilename, 0)
                    $Label.PrintOut(1, 0)
                    $PrintJobs += 1
                }
                catch {
                    Write-Output 'Failed'
                    Write-Output $Label.ErrorCode
                }
            }

            $Label.Close()
            $Label.EndPrint()
        }
        else {
            Write-Output 'Failed to open the label file'
        }
    }
    else {
        Write-Output 'Failed to find the specified printer'
    }
    if ($PrintJobs.Count -gt 0) {
    #    show the print jobs
        $PrintJobs | Format-Table
    }
    else {
        Write-Output 'No print jobs created'
    }
}
else {
    Write-Error "Error: $StatusCode"
    return
}
