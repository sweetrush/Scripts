# Function to get WiFi passwords
function Get-WifiPasswords {
   # Get all WiFi profiles
   $profiles = (netsh wlan show profiles) | Select-String "\:(.+)$" | ForEach-Object {
       $name = $_.Matches.Groups[1].Value.Trim()
       
       # Get password for each profile
       $password = (netsh wlan show profile name="$name" key=clear) |
           Select-String "Key Content\W+\:(.+)$" | ForEach-Object {
               $pass = $_.Matches.Groups[1].Value.Trim()
               if($pass) { $pass } else { "No Password Set" }
           }

       # Create custom object with network details
       [PSCustomObject]@{
           "Network" = $name
           "Password" = $password
           "Security" = (netsh wlan show profile name="$name" key=clear) |
               Select-String "Authentication\W+\:(.+)$" | ForEach-Object {
                   $_.Matches.Groups[1].Value.Trim()
               }
       }
   }

   # Export to JSON file
   $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
   $profiles | ConvertTo-Json | Out-File "wifi_passwords_$timestamp.json"

   # Display results
   $profiles | Format-Table -AutoSize
}

# Run the function
Get-WifiPasswords