#!/usr/bin/env powershell

<#
.SYNOPSIS
    Converts VBScript (.vbs) files to PowerShell (.ps1) scripts
.DESCRIPTION
    This script reads VBScript files and converts common VBScript constructs 
    to equivalent PowerShell syntax. It handles variables, functions, loops, 
    conditionals, and common VBScript objects.
.PARAMETER InputFile
    Path to the VBScript file to convert
.PARAMETER OutputFile
    Path for the output PowerShell file (optional)
.PARAMETER ShowProgress
    Display conversion progress and mappings
.EXAMPLE
    .\Convert-VBSToPS1.ps1 -InputFile "script.vbs" -OutputFile "script.ps1"
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$InputFile,
    
    [Parameter(Mandatory=$false)]
    [string]$OutputFile,
    
    [Parameter(Mandatory=$false)]
    [switch]$ShowProgress
)

function Convert-VBSToPS1 {
    param(
        [string]$VBSContent,
        [bool]$ShowProgress = $false
    )
    
    # Initialize conversion tracking
    $conversions = @()
    
    # Split content into lines for processing
    $lines = $VBSContent -split "`r?`n"
    $convertedLines = @()
    
    # VBScript to PowerShell mapping tables
    $variablePatterns = @{
        'Dim\s+(\w+)' = 'var $1'
        'Set\s+(\w+)\s*=\s*(.+)' = '$1 = $2'
        'Const\s+(\w+)\s*=\s*(.+)' = 'Set-Variable -Name $1 -Value $2 -Option Constant'
    }
    
    $objectPatterns = @{
        'CreateObject\("WScript\.Shell"\)' = 'New-Object -ComObject WScript.Shell'
        'CreateObject\("Scripting\.FileSystemObject"\)' = 'New-Object -ComObject Scripting.FileSystemObject'
        'CreateObject\("Excel\.Application"\)' = 'New-Object -ComObject Excel.Application'
        'CreateObject\("Word\.Application"\)' = 'New-Object -ComObject Word.Application'
        'CreateObject\("ADODB\.Connection"\)' = 'New-Object -ComObject ADODB.Connection'
        'CreateObject\("([^"]+)"\)' = 'New-Object -ComObject "$1"'
    }
    
    $functionPatterns = @{
        'Function\s+(\w+)\s*\((.*?)\)' = 'function $1($2) {'
        'Sub\s+(\w+)\s*\((.*?)\)' = 'function $1($2) {'
        'End\s+Function' = '}'
        'End\s+Sub' = '}'
    }
    
    $controlPatterns = @{
        'If\s+(.+?)\s+Then' = 'if ($1) {'
        'ElseIf\s+(.+?)\s+Then' = '} elseif ($1) {'
        'Else' = '} else {'
        'End\s+If' = '}'
        'For\s+(\w+)\s*=\s*(.+?)\s+To\s+(.+?)(?:\s+Step\s+(.+))?' = 'for ($1 = $2; $1 -le $3; $1 += $(if("$4") { $4 } else { 1 })) {'
        'For\s+Each\s+(\w+)\s+In\s+(.+)' = 'foreach ($1 in $2) {'
        'Next(?:\s+\w+)?' = '}'
        'While\s+(.+)' = 'while ($1) {'
        'Wend' = '}'
        'Do\s+While\s+(.+)' = 'do {'
        'Do\s+Until\s+(.+)' = 'do {'
        'Loop' = '} while ($condition)'
    }
    
    $operatorPatterns = @{
        '\bAnd\b' = '-and'
        '\bOr\b' = '-or'
        '\bNot\b' = '-not'
        '\bMod\b' = '%'
        '&' = '+'  # String concatenation
        '<>' = '-ne'
        '=' = '-eq'  # In conditional contexts
    }
    
    $methodPatterns = @{
        '\.Write\s*\(' = '.Write('
        '\.WriteLine\s*\(' = '.WriteLine('
        '\.MsgBox\s*\(' = '[System.Windows.Forms.MessageBox]::Show('
        'WScript\.Echo\s+(.+)' = 'Write-Host $1'
        'WScript\.Sleep\s+(\d+)' = 'Start-Sleep -Milliseconds $1'
        'WScript\.Quit(?:\s*\((\d+)\))?' = 'exit $(if("$1") { $1 } else { 0 })'
    }
    
    # Process each line
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $line = $lines[$i].Trim()
        $originalLine = $line
        
        # Skip empty lines and comments
        if ($line -eq '' -or $line.StartsWith("'")) {
            $convertedLines += $line
            continue
        }
        
        # Convert comments
        if ($line.Contains("'")) {
            $commentIndex = $line.IndexOf("'")
            $codePart = $line.Substring(0, $commentIndex).Trim()
            $commentPart = $line.Substring($commentIndex + 1)
            $line = $codePart + " # " + $commentPart
        }
        
        # Apply variable patterns
        foreach ($pattern in $variablePatterns.Keys) {
            if ($line -match $pattern) {
                $replacement = $variablePatterns[$pattern]
                $line = $line -replace $pattern, $replacement
                if ($ShowProgress) { $conversions += "Variable: $originalLine -> $line" }
                break
            }
        }
        
        # Apply object patterns
        foreach ($pattern in $objectPatterns.Keys) {
            if ($line -match $pattern) {
                $replacement = $objectPatterns[$pattern]
                $line = $line -replace $pattern, $replacement
                if ($ShowProgress) { $conversions += "Object: $originalLine -> $line" }
                break
            }
        }
        
        # Apply function patterns
        foreach ($pattern in $functionPatterns.Keys) {
            if ($line -match $pattern) {
                $replacement = $functionPatterns[$pattern]
                $line = $line -replace $pattern, $replacement
                if ($ShowProgress) { $conversions += "Function: $originalLine -> $line" }
                break
            }
        }
        
        # Apply control structure patterns
        foreach ($pattern in $controlPatterns.Keys) {
            if ($line -match $pattern) {
                $replacement = $controlPatterns[$pattern]
                $line = $line -replace $pattern, $replacement
                if ($ShowProgress) { $conversions += "Control: $originalLine -> $line" }
                break
            }
        }
        
        # Apply method patterns
        foreach ($pattern in $methodPatterns.Keys) {
            if ($line -match $pattern) {
                $replacement = $methodPatterns[$pattern]
                $line = $line -replace $pattern, $replacement
                if ($ShowProgress) { $conversions += "Method: $originalLine -> $line" }
                break
            }
        }
        
        # Apply operator patterns
        foreach ($pattern in $operatorPatterns.Keys) {
            if ($line -match $pattern) {
                $replacement = $operatorPatterns[$pattern]
                $line = $line -replace $pattern, $replacement
                if ($ShowProgress) { $conversions += "Operator: $originalLine -> $line" }
            }
        }
        
        # Handle variable assignments (remove Set keyword)
        $line = $line -replace '^Set\s+', ''
        
        # Handle string concatenation
        $line = $line -replace '([^+])\s*&\s*([^+])', '$1 + $2'
        
        # Handle VBScript string functions
        $line = $line -replace 'Len\(([^)]+)\)', '$1.Length'
        $line = $line -replace 'UCase\(([^)]+)\)', '$1.ToUpper()'
        $line = $line -replace 'LCase\(([^)]+)\)', '$1.ToLower()'
        $line = $line -replace 'Trim\(([^)]+)\)', '$1.Trim()'
        $line = $line -replace 'Left\(([^,]+),\s*([^)]+)\)', '$1.Substring(0, $2)'
        $line = $line -replace 'Right\(([^,]+),\s*([^)]+)\)', '$1.Substring($1.Length - $2)'
        $line = $line -replace 'Mid\(([^,]+),\s*([^,]+),\s*([^)]+)\)', '$1.Substring($2 - 1, $3)'
        $line = $line -replace 'InStr\(([^,]+),\s*([^)]+)\)', '$1.IndexOf($2) + 1'
        $line = $line -replace 'Replace\(([^,]+),\s*([^,]+),\s*([^)]+)\)', '$1.Replace($2, $3)'
        
        # Handle common VBScript constants
        $line = $line -replace '\bvbCrLf\b', '"`r`n"'
        $line = $line -replace '\bvbCr\b', '"`r"'
        $line = $line -replace '\bvbLf\b', '"`n"'
        $line = $line -replace '\bvbTab\b', '"`t"'
        $line = $line -replace '\bvbNullString\b', '""'
        
        # Add PowerShell variable prefix ($) if missing
        $line = $line -replace '\b(?<![$])\b([a-zA-Z_][a-zA-Z0-9_]*)\s*=', '$$1 ='
        
        $convertedLines += $line
    }
    
    # Add PowerShell header
    $header = @"
# Converted from VBScript to PowerShell
# Original file: $InputFile
# Conversion date: $(Get-Date)
# Note: This is an automated conversion. Manual review and testing recommended.

"@
    
    $result = $header + ($convertedLines -join "`n")
    
    if ($ShowProgress) {
        Write-Host "Conversion Summary:" -ForegroundColor Green
        $conversions | ForEach-Object { Write-Host "  $_" -ForegroundColor Yellow }
    }
    
    return $result
}

# Main execution
try {
    if (-not (Test-Path $InputFile)) {
        throw "Input file '$InputFile' not found."
    }
    
    Write-Host "Reading VBScript file: $InputFile" -ForegroundColor Cyan
    $vbsContent = Get-Content $InputFile -Raw
    
    Write-Host "Converting VBScript to PowerShell..." -ForegroundColor Cyan
    $psContent = Convert-VBSToPS1 -VBSContent $vbsContent -ShowProgress $ShowProgress
    
    # Determine output file name
    if (-not $OutputFile) {
        $OutputFile = [System.IO.Path]::ChangeExtension($InputFile, '.ps1')
    }
    
    Write-Host "Writing PowerShell file: $OutputFile" -ForegroundColor Cyan
    $psContent | Out-File -FilePath $OutputFile -Encoding UTF8
    
    Write-Host "Conversion completed successfully!" -ForegroundColor Green
    Write-Host "Output file: $OutputFile" -ForegroundColor Green
    
    # Display warnings
    Write-Host "`nIMPORTANT NOTES:" -ForegroundColor Yellow
    Write-Host "1. This is an automated conversion - manual review is required" -ForegroundColor Yellow
    Write-Host "2. Test the converted script thoroughly before using in production" -ForegroundColor Yellow
    Write-Host "3. Some VBScript features may not have direct PowerShell equivalents" -ForegroundColor Yellow
    Write-Host "4. Error handling and COM object cleanup may need adjustment" -ForegroundColor Yellow
    
} catch {
    Write-Error "Conversion failed: $($_.Exception.Message)"
    exit 1
}
