# ===================== CONFIGURAÇÃO =====================

$log = "C:\temp\tape_expired_remove.log"

# ===================== INICIALIZAÇÃO =====================

Import-Module Veeam.Backup.PowerShell -ErrorAction Stop

Get-Date | Out-File $log
Write-Output "===== INÍCIO DA EXECUÇÃO =====" | Out-File $log -Append

$now = Get-Date
$RemovedTapes = @()

# Obter todos os Media Pools
$Pools = Get-VBRTapeMediaPool

if (-not $Pools) {
    Write-Host "Nenhum Media Pool encontrado."
    Write-Output "Nenhum Media Pool encontrado." | Out-File $log -Append
    exit
}

# ===================== LOOP POR POOL =====================

foreach ($pool in $Pools) {

    Write-Host ""
    $answer = Read-Host "Deseja verificar o pool '$($pool.Name)'? (S/N)"
    Write-Output "Pool: $($pool.Name) | Verificar: $answer" | Out-File $log -Append

    if ($answer -notmatch "^[sS]$") {
        continue
    }

    $ExpiredTapes = Get-VBRTapeMedium -MediaPool $pool | Where-Object {
        $_.ProtectedBySoftware -eq $true -and
        $_.ExpirationDate -ne $null -and
        $_.ExpirationDate -le $now
    }

    if (-not $ExpiredTapes) {
        Write-Host "Nenhuma fita expirada no pool $($pool.Name)."
        Write-Output "Nenhuma fita expirada no pool $($pool.Name)." | Out-File $log -Append
        continue
    }

    # Mostrar detalhes antes da confirmação
    Write-Host ""
    Write-Host "Fitas expiradas encontradas no pool '$($pool.Name)':"
    Write-Host ""

    $ExpiredTapes |
        Select-Object Name, MediaSet, ExpirationDate |
        Sort-Object ExpirationDate |
        Format-Table -AutoSize

    Write-Output "Fitas expiradas no pool $($pool.Name):" | Out-File $log -Append
    $ExpiredTapes |
        Select-Object Name, MediaSet, ExpirationDate |
        Format-Table -AutoSize | Out-File $log -Append

    Write-Host ""
    $confirm = Read-Host "Deseja DESPROTEGER e REMOVER essas fitas do pool '$($pool.Name)'? (S/N)"
    Write-Output "Confirmação remoção pool $($pool.Name): $confirm" | Out-File $log -Append

    if ($confirm -notmatch "^[sS]$") {
        Write-Host "Operação cancelada para este pool."
        Write-Output "Remoção cancelada para o pool $($pool.Name)." | Out-File $log -Append
        continue
    }

    foreach ($t in $ExpiredTapes) {
        try {
            Disable-VBRTapeProtection -Medium $t
            Remove-VBRTapeMedium -Medium $t -Confirm:$false

            $RemovedTapes += $t.Name
            Write-Output "SUCESSO: $($t.Name) removida do pool $($pool.Name)" | Out-File $log -Append
        }
        catch {
            Write-Output "ERRO: $($t.Name) | $_" | Out-File $log -Append
        }
    }
}

# ===================== RESULTADO FINAL =====================

Write-Host ""
Write-Host "===== FITAS DESPROTEGIDAS E REMOVIDAS ====="

Write-Output "===== RESUMO FINAL =====" | Out-File $log -Append

if ($RemovedTapes.Count -eq 0) {
    Write-Host "Nenhuma fita foi removida."
    Write-Output "Nenhuma fita foi removida." | Out-File $log -Append
} else {
    $RemovedTapes | Sort-Object | ForEach-Object {
        Write-Host $_
        Write-Output $_ | Out-File $log -Append
    }
}

Write-Output "===== FIM DA EXECUÇÃO =====" | Out-File $log -Append
