Remove-Item number.db
Import-Module PSLiteDB
Open-LiteDBConnection .\number.db
New-LiteDBCollection batchinformation



function JobReport () {



    # Build PScustomObject
    $JobReport = @()
    $JobReport = [pscustomobject][ordered]@{}
    $JobReport | Add-Member -MemberType NoteProperty -Name "_id" -Value "0001"
    $JobReport | Add-Member -MemberType NoteProperty -Name "Date"-Value "something"
    $JobReport | ConvertTo-LiteDbBSON | Add-LiteDBDocument -Collection batchinformation

    $JobReport | Add-Member -MemberType NoteProperty -Name "_id" -Value "0002"
    $JobReport | Add-Member -MemberType NoteProperty -Name "Date"-Value "something"
    $JobReport | ConvertTo-LiteDbBSON | Add-LiteDBDocument -Collection batchinformation

    $JobReport | Add-Member -MemberType NoteProperty -Name "_id" -Value "0003"
    $JobReport | Add-Member -MemberType NoteProperty -Name "Date"-Value "something"
    $JobReport | ConvertTo-LiteDbBSON | Add-LiteDBDocument -Collection batchinformation



}


# function run order
JobReport

