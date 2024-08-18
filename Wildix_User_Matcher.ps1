# Prompt the admin for the paths to the source file and the target file
$source_csv = Read-Host "Please enter the path to the source file (CSV)"
$target_csv = Read-Host "Please enter the path to the target file (CSV)"
$updated_target_csv = Read-Host "Please enter the path for the updated target file (CSV) (default: updated_target.csv)"
if (-not $updated_target_csv) {
    $updated_target_csv = "updated_target.csv"
}

# Import the CSV files
$source = Import-Csv -Path $source_csv
$target = Import-Csv -Path $target_csv

# Array to store the updated rows of the target file
$updated_target = @()

# Iterate through the rows of the target file
foreach ($target_row in $target) {
    # Find the matching row in the source file based on UPN and Email
    $matching_row = $source | Where-Object { $_.UPN -eq $target_row."Email" }
    
    # If a matching row is found
    if ($matching_row) {
        # Clear the existing value in the ImageURL column
        $target_row."ImageURL" = ""
        
        # Insert the new value from the URL column of the source file
        $target_row."ImageURL" = $matching_row.URL
    }

    # Add the (possibly updated) row to the new list
    $updated_target += $target_row
}

# Save the updated target file
$updated_target | Export-Csv -Path $updated_target_csv -NoTypeInformation

Write-Host "The target file has been successfully updated and saved to '$updated_target_csv'."
