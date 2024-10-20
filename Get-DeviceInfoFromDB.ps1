Add-Type -AssemblyName System.Windows.Forms

$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{
    InitialDirectory = $PSScriptRoot
    Filter = 'Text Documents (*.txt)|*.txt'
}

$result = $FileBrowser.ShowDialog()

if ($result -eq "Cancel") {
    exit
}

$computerNames = Get-Content -Path $FileBrowser.FileName

# Set the DSN name
$dsnName = "YourDBName"

# Create a new ODBC connection
$conn = New-Object System.Data.Odbc.OdbcConnection("DSN=$dsnName")

# Open the connection
$conn.Open()

# Get system information for each computer
$computerNames | ForEach-Object {
    try {
            # Set the SQL Query
            $SQLQuery = @("SELECT [Device],[HostName],[Type],[Manufacturer],[Description],[Modell],[SerialNumber],[ServiceInfo],[Location1],[Location2],[Location3],[Department],[EndUser],[AssetNumber],[ScreenSize],[MacAddress],[Criticality] FROM [MyDevices].[dbo].[DevicesALL] WHERE Device = '$_'")
                        
            # Create a new ODBC command
            $cmd = New-Object System.Data.Odbc.OdbcCommand($sqlQuery, $conn)
            # Execute the command and store the results in a data table
            $dataTable = New-Object System.Data.DataTable
            $dataTable.Load($cmd.ExecuteReader())

            $GridView += $dataTable

        } catch {
            Write-Host $_ -ForegroundColor Red
    }
}

# Close the connection
$conn.Close()

# Display the information in an interactive grid view
$GridView | Select-Object Device, HostName, Type, Manufacturer, Description, Modell, SerialNumber, ServiceInfo, Location1, Location2, Location3, Department, EndUser, AssetNumber, ScreenSize, MacAddress, Criticality | Out-GridView -Wait -Title "GridView"
