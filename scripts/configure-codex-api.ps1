$ErrorActionPreference = "Stop"

if (-not (Get-Command codex -ErrorAction SilentlyContinue)) {
    throw "Codex CLI is not available on PATH."
}

$secureKey = Read-Host "Enter your OpenAI API key" -AsSecureString
$keyPointer = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureKey)

try {
    $plainKey = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($keyPointer)
    $plainKey | & codex login --with-api-key
    if ($LASTEXITCODE -ne 0) {
        throw "Codex API-key login failed with exit code $LASTEXITCODE."
    }
}
finally {
    if ($keyPointer -ne [IntPtr]::Zero) {
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($keyPointer)
    }
    Remove-Variable plainKey -ErrorAction SilentlyContinue
}

& codex login status
