# ==============================================================================
# Monitor de Latência DNS (BETA V3)
# Criador: Fabiopsyduck
# ==============================================================================

# ==============================================================================
# 0. TRAVA DE INSTÂNCIA ÚNICA (MUTEX)
# ==============================================================================
$mutexCreated = $false
# Cria uma assinatura única no núcleo do Windows para o seu programa
$global:appMutex = New-Object System.Threading.Mutex($true, "Global\MonitorLatenciaDNS_Fabiopsyduck", [ref]$mutexCreated)

if (-not $mutexCreated) {
    # Se a assinatura já existir, outra janela está aberta. Mostra o aviso e mata o processo.
    Add-Type -AssemblyName System.Windows.Forms
    [void][System.Windows.Forms.MessageBox]::Show("O Monitor de Latência DNS já está em execução no seu computador.`n`nVerifique a sua barra de tarefas.", "Instância Duplicada", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    exit
}

# ==============================================================================
# 1. GESTÃO DE MEMÓRIA GDI (ÍCONES) E BIBLIOTECAS
# ==============================================================================
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public static class Win32 {
    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool DestroyIcon(IntPtr hIcon);
}
"@

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms.DataVisualization

# ==============================================================================
# 2. GESTÃO DE ARQUIVOS E PASTAS (CSV) E VARIÁVEIS COMPARTILHADAS
# ==============================================================================
# Blindagem de diretório raiz (Compatibilidade perfeita entre .ps1 e compiladores .exe)
$processPath = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName

if ($processPath -match "(powershell\.exe|pwsh\.exe|powershell_ise\.exe)$") {
    # Se for o PowerShell rodando, usamos a raiz do script
    $scriptPath = $PSScriptRoot
} else {
    # Se for o .exe já compilado, pegamos a pasta exata de onde o usuário clicou
    $scriptPath = Split-Path -Path $processPath
}

# Trava de segurança final
if ([string]::IsNullOrEmpty($scriptPath)) { $scriptPath = $PWD.Path }

$dnsFolder = Join-Path $scriptPath "Servidores_DNS"
$fileIPv4 = Join-Path $dnsFolder "lista_ipv4.csv"
$fileIPv6 = Join-Path $dnsFolder "lista_ipv6.csv"

# --- Caminho do arquivo da Lista Negra ---
$fileBlacklist = Join-Path $dnsFolder "lista_negra.csv" 

if (-not (Test-Path $dnsFolder)) { New-Item -ItemType Directory -Path $dnsFolder | Out-Null }

function Test-OrInitializeCSV ($filePath, $isIPv4) {
    if (-not (Test-Path $filePath)) {
        $header = "Nome,DNS_Primario,DNS_Secundario,Selecionado"
        $header | Out-File -FilePath $filePath -Encoding UTF8
        if ($isIPv4) {
            "Google,8.8.8.8,8.8.4.4,True" | Out-File -FilePath $filePath -Encoding UTF8 -Append
            "Cloudflare,1.1.1.1,1.0.0.1,True" | Out-File -FilePath $filePath -Encoding UTF8 -Append
        } else {
            "Google,2001:4860:4860::8888,2001:4860:4860::8844,True" | Out-File -FilePath $filePath -Encoding UTF8 -Append
        }
    }
}

# --- Inicializador da Lista Negra ---
function Test-OrInitializeBlacklist {
    if (-not (Test-Path $fileBlacklist)) {
        "IP,Nome,Status" | Out-File -FilePath $fileBlacklist -Encoding UTF8
    }
}

Test-OrInitializeCSV -filePath $fileIPv4 -isIPv4 $true
Test-OrInitializeCSV -filePath $fileIPv6 -isIPv4 $false
Test-OrInitializeBlacklist

# Variáveis globais para o Multithread
$global:psInst = $null
$global:asyncResult = $null
$global:rankingText = "" 
$global:syncHash = [hashtable]::Synchronized(@{
    MacroCount = 0; TotalMacro = 1; MacroName = "";
    MicroCount = 0; TotalMicro = 100;
    Cancel = $false; IsRunning = $false;
    Resultados = [System.Collections.ArrayList]::new()
})

# ==============================================================================
# 3. JANELA CUSTOMIZADA PARA ADICIONAR / EDITAR DNS (BLINDADA CONTRA REPETIÇÕES E TIPO DE IP)
# ==============================================================================
function Show-DNSDialog {
    param (
        [string]$Titulo = "Novo DNS", 
        [string]$NomeAtual = "", 
        [string]$PrimarioAtual = "", 
        [string]$SecundarioAtual = "",
        [array]$ListaExistente = @(),
        [string]$TipoIP = "IPv4"
    )
    
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = $Titulo; $dlg.Size = New-Object System.Drawing.Size(350, 270)
    $dlg.StartPosition = "CenterParent"; $dlg.FormBorderStyle = "FixedDialog"
    $dlg.MaximizeBox = $false; $dlg.MinimizeBox = $false
    $dlg.BackColor = "White"

    $fonteLabel = New-Object System.Drawing.Font("Segoe UI", 10)
    
    $ctxMenu = New-Object System.Windows.Forms.ContextMenuStrip
    $itemRecortar = $ctxMenu.Items.Add("Recortar")
    $itemCopiar = $ctxMenu.Items.Add("Copiar")
    $itemColar = $ctxMenu.Items.Add("Colar")
    $ctxMenu.Items.Add("-") | Out-Null
    $itemSelTudo = $ctxMenu.Items.Add("Selecionar Tudo")

    $itemRecortar.Add_Click({ if ($ctxMenu.SourceControl) { $ctxMenu.SourceControl.Cut() } })
    $itemCopiar.Add_Click({ if ($ctxMenu.SourceControl) { $ctxMenu.SourceControl.Copy() } })
    $itemColar.Add_Click({ if ($ctxMenu.SourceControl) { $ctxMenu.SourceControl.Paste() } })
    $itemSelTudo.Add_Click({ if ($ctxMenu.SourceControl) { $ctxMenu.SourceControl.SelectAll() } })
    
    $lblNome = New-Object System.Windows.Forms.Label; $lblNome.Text = "Nome:"; $lblNome.Location = New-Object System.Drawing.Point(20, 20); $lblNome.AutoSize = $true; $lblNome.Font = $fonteLabel
    $txtNome = New-Object System.Windows.Forms.TextBox; $txtNome.Location = New-Object System.Drawing.Point(20, 45); $txtNome.Size = New-Object System.Drawing.Size(290, 25); $txtNome.Text = $NomeAtual; $txtNome.Font = $fonteLabel
    $txtNome.ContextMenuStrip = $ctxMenu 
    
    $lblPri = New-Object System.Windows.Forms.Label; $lblPri.Text = "IP Primário:"; $lblPri.Location = New-Object System.Drawing.Point(20, 75); $lblPri.AutoSize = $true; $lblPri.Font = $fonteLabel
    $txtPri = New-Object System.Windows.Forms.TextBox; $txtPri.Location = New-Object System.Drawing.Point(20, 100); $txtPri.Size = New-Object System.Drawing.Size(290, 25); $txtPri.Text = $PrimarioAtual; $txtPri.Font = $fonteLabel
    $txtPri.ContextMenuStrip = $ctxMenu 
    
    $lblSec = New-Object System.Windows.Forms.Label; $lblSec.Text = "IP Secundário:"; $lblSec.Location = New-Object System.Drawing.Point(20, 130); $lblSec.AutoSize = $true; $lblSec.Font = $fonteLabel
    $txtSec = New-Object System.Windows.Forms.TextBox; $txtSec.Location = New-Object System.Drawing.Point(20, 155); $txtSec.Size = New-Object System.Drawing.Size(290, 25); $txtSec.Text = $SecundarioAtual; $txtSec.Font = $fonteLabel
    $txtSec.ContextMenuStrip = $ctxMenu

    $btnSalvar = New-Object System.Windows.Forms.Button; $btnSalvar.Text = "Salvar"; $btnSalvar.Location = New-Object System.Drawing.Point(120, 195); $btnSalvar.Size = New-Object System.Drawing.Size(90, 30)
    $btnSalvar.FlatStyle = "Flat"; $btnSalvar.FlatAppearance.BorderSize = 0; $btnSalvar.BackColor = [System.Drawing.Color]::SteelBlue; $btnSalvar.ForeColor = "White"; $btnSalvar.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold); $btnSalvar.Cursor = [System.Windows.Forms.Cursors]::Hand

    $btnCancelar = New-Object System.Windows.Forms.Button; $btnCancelar.Text = "Cancelar"; $btnCancelar.Location = New-Object System.Drawing.Point(220, 195); $btnCancelar.Size = New-Object System.Drawing.Size(90, 30); $btnCancelar.DialogResult = "Cancel"
    $btnCancelar.FlatStyle = "Flat"; $btnCancelar.FlatAppearance.BorderSize = 0; $btnCancelar.BackColor = [System.Drawing.Color]::SlateGray; $btnCancelar.ForeColor = "White"; $btnCancelar.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold); $btnCancelar.Cursor = [System.Windows.Forms.Cursors]::Hand

    $btnSalvar.Add_Click({
        $novoNome = $txtNome.Text.Trim()
        $novoPri = $txtPri.Text.Trim()
        $novoSec = $txtSec.Text.Trim()

        if ([string]::IsNullOrWhiteSpace($novoNome)) {
            [System.Windows.Forms.MessageBox]::Show("O nome do DNS não pode estar vazio.", "Aviso", 0, "Warning")
            return
        }

        # ==========================================================================
        # 1. VALIDADOR DE FAMÍLIA DE IP
        # ==========================================================================
        $familiaEsperada = if ($TipoIP -eq "IPv6") { 'InterNetworkV6' } else { 'InterNetwork' }
        $ipObj = $null

        if (-not [string]::IsNullOrWhiteSpace($novoPri)) {
            if (-not [System.Net.IPAddress]::TryParse($novoPri, [ref]$ipObj)) {
                [System.Windows.Forms.MessageBox]::Show("O IP Primário '$novoPri' não tem um formato válido.", "Erro de Digitação", 0, "Error")
                return
            }
            if ($ipObj.AddressFamily.ToString() -ne $familiaEsperada) {
                [System.Windows.Forms.MessageBox]::Show("Você está na lista de $TipoIP, mas o IP Primário inserido pertence à família oposta.", "Conflito de IP", 0, "Warning")
                return
            }
        }

        if (-not [string]::IsNullOrWhiteSpace($novoSec)) {
            if (-not [System.Net.IPAddress]::TryParse($novoSec, [ref]$ipObj)) {
                [System.Windows.Forms.MessageBox]::Show("O IP Secundário '$novoSec' não tem um formato válido.", "Erro de Digitação", 0, "Error")
                return
            }
            if ($ipObj.AddressFamily.ToString() -ne $familiaEsperada) {
                [System.Windows.Forms.MessageBox]::Show("Você está na lista de $TipoIP, mas o IP Secundário inserido pertence à família oposta.", "Conflito de IP", 0, "Warning")
                return
            }
        }

        # ==========================================================================
        # 2. MOTOR FISCAL DE REPETIÇÕES
        # ==========================================================================
        $duplicado = $false
        $msgErro = ""

        foreach ($item in $ListaExistente) {
            if (-not [string]::IsNullOrEmpty($NomeAtual) -and $item.Nome -eq $NomeAtual) { continue }

            if ($item.Nome.Trim() -eq $novoNome) {
                $msgErro = "Já existe um servidor cadastrado com o nome '$novoNome'."; $duplicado = $true; break
            }
            if (-not [string]::IsNullOrWhiteSpace($novoPri)) {
                if ($item.DNS_Primario -eq $novoPri -or $item.DNS_Secundario -eq $novoPri) {
                    $msgErro = "O IP '$novoPri' já está em uso pelo servidor '$($item.Nome)'."; $duplicado = $true; break
                }
            }
            if (-not [string]::IsNullOrWhiteSpace($novoSec)) {
                if ($item.DNS_Primario -eq $novoSec -or $item.DNS_Secundario -eq $novoSec) {
                    $msgErro = "O IP '$novoSec' já está em uso pelo servidor '$($item.Nome)'."; $duplicado = $true; break
                }
            }
        }

        if ($duplicado) {
            [System.Windows.Forms.MessageBox]::Show($msgErro, "Entrada Duplicada", 0, "Warning")
            return
        }

        $dlg.DialogResult = "OK"
        $dlg.Close()
    })

    $dlg.Controls.Add($lblNome); $dlg.Controls.Add($txtNome); $dlg.Controls.Add($lblPri); $dlg.Controls.Add($txtPri); $dlg.Controls.Add($lblSec); $dlg.Controls.Add($txtSec)
    $dlg.Controls.Add($btnSalvar); $dlg.Controls.Add($btnCancelar)
    $dlg.AcceptButton = $btnSalvar; $dlg.CancelButton = $btnCancelar

    if ($dlg.ShowDialog() -eq "OK") {
        $retorno = [PSCustomObject]@{ Nome = $txtNome.Text.Trim(); DNS_Primario = $txtPri.Text.Trim(); DNS_Secundario = $txtSec.Text.Trim() }
        $dlg.Dispose(); return $retorno
    }
    $dlg.Dispose(); return $null
}

# ==============================================================================
# 4. CONSTRUÇÃO DA INTERFACE GRÁFICA (WINFORMS)
# ==============================================================================

# ==============================================================================
# GERADOR DE ÍCONE DPI-AWARE (Sem Memory Leak)
# ==============================================================================
$global:hIconGlobe = $null

function Get-DPIAwareIcon {
    # Tamanho 64x64 garante nitidez extrema em telas grandes
    $bmp = New-Object System.Drawing.Bitmap(64, 64)
    $gfx = [System.Drawing.Graphics]::FromImage($bmp)
    
    # Aplica Anti-Aliasing para as bordas do ícone ficarem redondinhas
    $gfx.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit
    $gfx.Clear([System.Drawing.Color]::Transparent)
    
    # Desenha o ícone na cor do tema do programa
    $brush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::SteelBlue)
    $font = New-Object System.Drawing.Font("Segoe MDL2 Assets", 46, [System.Drawing.GraphicsUnit]::Pixel)
    
    $fmt = New-Object System.Drawing.StringFormat
    $fmt.Alignment = [System.Drawing.StringAlignment]::Center
    $fmt.LineAlignment = [System.Drawing.StringAlignment]::Center
    
    $rect = New-Object System.Drawing.RectangleF(0, 0, 64, 64)
    $gfx.DrawString([char]0xE9D2, $font, $brush, $rect, $fmt)
    
    # Converte a imagem vetorial para Handle de Ícone do Windows
    $global:hIconGlobe = $bmp.GetHicon()
    $icon = [System.Drawing.Icon]::FromHandle($global:hIconGlobe)
    
    # Descarta as ferramentas de desenho
    $gfx.Dispose(); $brush.Dispose(); $font.Dispose(); $fmt.Dispose(); $bmp.Dispose()
    
    return $icon
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "Monitor de Latência DNS  (Criador: Fabiopsyduck) (Versão: BETA V3)"
$form.Icon = Get-DPIAwareIcon
$form.ClientSize = New-Object System.Drawing.Size(850, 600)
$form.StartPosition = "CenterScreen"
$form.MinimumSize = New-Object System.Drawing.Size(800, 550)
$form.BackColor = "White"

# --- PAINEL DE AVISO (PREPARAÇÃO) ---
$pnlWarning = New-Object System.Windows.Forms.Panel
$pnlWarning.Dock = "Fill"
$pnlWarning.BackColor = "White"

$pnlCard = New-Object System.Windows.Forms.Panel
$pnlCard.AutoSize = $true
$pnlCard.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
$pnlCard.BackColor = "White"

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = "Preparação para o Teste DNS"
$lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 18, [System.Drawing.FontStyle]::Bold)
$lblTitle.ForeColor = [System.Drawing.Color]::SteelBlue
$lblTitle.AutoSize = $true
$lblTitle.Location = New-Object System.Drawing.Point(0, 10)

$lblSubTitle = New-Object System.Windows.Forms.Label
$lblSubTitle.Text = "PARA RESULTADOS MAIS PRECISOS, É RECOMENDADO:"
$lblSubTitle.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$lblSubTitle.ForeColor = [System.Drawing.Color]::DimGray
$lblSubTitle.AutoSize = $true
$lblSubTitle.Location = New-Object System.Drawing.Point(0, 60)

$lblBody = New-Object System.Windows.Forms.Label
$lblBody.Text = "1. Evite atividades online durante o teste:`n   - Navegação web`n   - Streaming (YouTube, Netflix)`n   - Downloads/uploads`n   - Clientes torrent (qBittorrent, uTorrent, BitTorrent)`n   - Jogos online (Steam, Epic Games)`n`n2. Otimize sua conexão:`n   - Use cabo de rede (evite Wi-Fi)`n   - Desative VPNs`n   - Desative antivírus temporariamente (opcional)`n   - Desative apps que podem alterar ou manipular DNS`n`n3. Reduza interferências:`n   - Evite programas pesados (jogos, edição)`n   - Certifique-se de que não há atualizações em andamento (Windows, Steam, Battle.net)"
$lblBody.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Regular)
$lblBody.ForeColor = [System.Drawing.Color]::Black
$lblBody.AutoSize = $true
$lblBody.Location = New-Object System.Drawing.Point(20, 95)
$lblBody.Padding = New-Object System.Windows.Forms.Padding(0, 0, 0, 20) 

$pnlCard.Controls.Add($lblTitle)
$pnlCard.Controls.Add($lblSubTitle)
$pnlCard.Controls.Add($lblBody)

$pnlWarningBot = New-Object System.Windows.Forms.Panel
$pnlWarningBot.Dock = "Bottom"
$pnlWarningBot.Height = 100
$pnlWarningBot.BackColor = "White"

$btnProsseguir = New-Object System.Windows.Forms.Button
$btnProsseguir.Text = "Prosseguir para o Teste"
$btnProsseguir.Size = New-Object System.Drawing.Size(240, 45)
$btnProsseguir.BackColor = [System.Drawing.Color]::SteelBlue
$btnProsseguir.ForeColor = [System.Drawing.Color]::White
$btnProsseguir.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnProsseguir.FlatAppearance.BorderSize = 0
$btnProsseguir.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
$btnProsseguir.Cursor = [System.Windows.Forms.Cursors]::Hand

$pnlWarningBot.Add_Resize({
    $btnProsseguir.Left = ($pnlWarningBot.Width - $btnProsseguir.Width) / 2
    $btnProsseguir.Top = 20
})

$pnlWarning.Add_Resize({
    $pnlCard.Left = ($pnlWarning.Width - $pnlCard.Width) / 2
    $pnlCard.Top = (($pnlWarning.Height - $pnlWarningBot.Height) - $pnlCard.Height) / 2
    if ($lblTitle.Width -lt $pnlCard.Width) { $lblTitle.Left = ($pnlCard.Width - $lblTitle.Width) / 2 }
    if ($lblSubTitle.Width -lt $pnlCard.Width) { $lblSubTitle.Left = ($pnlCard.Width - $lblSubTitle.Width) / 2 }
})

$pnlWarningBot.Controls.Add($btnProsseguir)
$pnlWarning.Controls.Add($pnlCard)
$pnlWarning.Controls.Add($pnlWarningBot)
$form.Controls.Add($pnlWarning)

# --- PAINEL PRINCIPAL (MENU DE DNS) ---
$pnlMain = New-Object System.Windows.Forms.Panel
$pnlMain.Dock = "Fill"
$pnlMain.Visible = $false

# 1. BARRA SUPERIOR (Falsas Abas e Botões de Marcar/Desmarcar)
$pnlTopBar = New-Object System.Windows.Forms.Panel
$pnlTopBar.Width = 850 
$pnlTopBar.Dock = "Top"
$pnlTopBar.Height = 40
$pnlTopBar.BackColor = [System.Drawing.Color]::WhiteSmoke

$btnTabIPv4 = New-Object System.Windows.Forms.Button
$btnTabIPv4.Text = "Servidores IPv4"; $btnTabIPv4.Size = New-Object System.Drawing.Size(140, 40); $btnTabIPv4.Location = New-Object System.Drawing.Point(0, 0)
$btnTabIPv4.FlatStyle = "Flat"; $btnTabIPv4.FlatAppearance.BorderSize = 0
# COR DA ABA SELECIONADA (IPv4 inicia selecionada)
$btnTabIPv4.BackColor = [System.Drawing.Color]::SteelBlue; $btnTabIPv4.ForeColor = "White"; $btnTabIPv4.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold); $btnTabIPv4.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnTabIPv4.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::SteelBlue
$btnTabIPv4.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::SteelBlue

$btnTabIPv6 = New-Object System.Windows.Forms.Button
$btnTabIPv6.Text = "Servidores IPv6"; $btnTabIPv6.Size = New-Object System.Drawing.Size(140, 40); $btnTabIPv6.Location = New-Object System.Drawing.Point(140, 0)
$btnTabIPv6.FlatStyle = "Flat"; $btnTabIPv6.FlatAppearance.BorderSize = 0
# COR DA ABA INATIVA
$btnTabIPv6.BackColor = [System.Drawing.Color]::LightGray; $btnTabIPv6.ForeColor = "Black"; $btnTabIPv6.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Regular); $btnTabIPv6.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnTabIPv6.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::LightGray
$btnTabIPv6.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::LightGray

$btnMarcarTopo = New-Object System.Windows.Forms.Button
$btnMarcarTopo.Text = "Marcar Todas"; $btnMarcarTopo.Size = New-Object System.Drawing.Size(110, 26)
$btnMarcarTopo.FlatStyle = "Flat"; $btnMarcarTopo.FlatAppearance.BorderSize = 0; $btnMarcarTopo.BackColor = [System.Drawing.Color]::DarkGray; $btnMarcarTopo.ForeColor = "White"; $btnMarcarTopo.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold); $btnMarcarTopo.Cursor = [System.Windows.Forms.Cursors]::Hand

$btnDesmarcarTopo = New-Object System.Windows.Forms.Button
$btnDesmarcarTopo.Text = "Desmarcar Todas"; $btnDesmarcarTopo.Size = New-Object System.Drawing.Size(125, 26)
$btnDesmarcarTopo.FlatStyle = "Flat"; $btnDesmarcarTopo.FlatAppearance.BorderSize = 0; $btnDesmarcarTopo.BackColor = [System.Drawing.Color]::DarkGray; $btnDesmarcarTopo.ForeColor = "White"; $btnDesmarcarTopo.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold); $btnDesmarcarTopo.Cursor = [System.Windows.Forms.Cursors]::Hand

$pnlTopBar.Controls.Add($btnTabIPv4); $pnlTopBar.Controls.Add($btnTabIPv6); $pnlTopBar.Controls.Add($btnMarcarTopo); $pnlTopBar.Controls.Add($btnDesmarcarTopo)
$pnlTopBar.Add_Resize({
    $btnDesmarcarTopo.Left = $pnlTopBar.Width - $btnDesmarcarTopo.Width - 15
    $btnDesmarcarTopo.Top = ($pnlTopBar.Height - $btnDesmarcarTopo.Height) / 2
    
    $btnMarcarTopo.Left = $btnDesmarcarTopo.Left - $btnMarcarTopo.Width - 10
    $btnMarcarTopo.Top = ($pnlTopBar.Height - $btnMarcarTopo.Height) / 2
})

# 2. CONTAINERS PARA RECEBER AS TABELAS
$pnlContentIPv4 = New-Object System.Windows.Forms.Panel; $pnlContentIPv4.Dock = "Fill"; $pnlContentIPv4.BackColor = "White"
$pnlContentIPv6 = New-Object System.Windows.Forms.Panel; $pnlContentIPv6.Dock = "Fill"; $pnlContentIPv6.BackColor = "White"; $pnlContentIPv6.Visible = $false

# Adiciona tudo ao painel principal
$pnlMain.Controls.Add($pnlContentIPv4)
$pnlMain.Controls.Add($pnlContentIPv6)
$pnlMain.Controls.Add($pnlTopBar)

# CORREÇÃO DE PROFUNDIDADE
$pnlTopBar.SendToBack()
$pnlContentIPv4.BringToFront()

# 3. LÓGICA DE NAVEGAÇÃO DAS FALSAS ABAS
$btnTabIPv4.Add_Click({
    $pnlContentIPv4.Visible = $true; $pnlContentIPv6.Visible = $false
    
    # Atualiza a aba IPv4 para a cor SELECIONADA
    $btnTabIPv4.BackColor = [System.Drawing.Color]::SteelBlue; $btnTabIPv4.ForeColor = "White"; $btnTabIPv4.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $btnTabIPv4.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::SteelBlue
    $btnTabIPv4.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::SteelBlue

    # Atualiza a aba IPv6 para a cor INATIVA
    $btnTabIPv6.BackColor = [System.Drawing.Color]::LightGray; $btnTabIPv6.ForeColor = "Black"; $btnTabIPv6.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Regular)
    $btnTabIPv6.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::LightGray
    $btnTabIPv6.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::LightGray
})
$btnTabIPv6.Add_Click({
    $pnlContentIPv6.Visible = $true; $pnlContentIPv4.Visible = $false
    
    # Atualiza a aba IPv6 para a cor SELECIONADA
    $btnTabIPv6.BackColor = [System.Drawing.Color]::SteelBlue; $btnTabIPv6.ForeColor = "White"; $btnTabIPv6.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $btnTabIPv6.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::SteelBlue
    $btnTabIPv6.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::SteelBlue

    # Atualiza a aba IPv4 para a cor INATIVA
    $btnTabIPv4.BackColor = [System.Drawing.Color]::LightGray; $btnTabIPv4.ForeColor = "Black"; $btnTabIPv4.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Regular)
    $btnTabIPv4.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::LightGray
    $btnTabIPv4.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::LightGray
})

$form.Controls.Add($pnlMain)

$btnProsseguir.Add_Click({
    $pnlWarning.Visible = $false
    $pnlMain.Visible = $true
})

# --- PAINEL DE EXECUÇÃO (TESTE EM ANDAMENTO) ---
$pnlExec = New-Object System.Windows.Forms.Panel
$pnlExec.Dock = "Fill"
$pnlExec.BackColor = "White"
$pnlExec.Visible = $false

$lblMacroExec = New-Object System.Windows.Forms.Label
$lblMacroExec.Text = "Servidor: Aguardando... (0/0)"
$lblMacroExec.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
$lblMacroExec.ForeColor = [System.Drawing.Color]::SteelBlue
$lblMacroExec.Location = New-Object System.Drawing.Point(50, 100)
$lblMacroExec.AutoSize = $true

$pbMacro = New-Object System.Windows.Forms.ProgressBar
$pbMacro.Location = New-Object System.Drawing.Point(50, 140)
$pbMacro.Height = 25

$lblMicroExec = New-Object System.Windows.Forms.Label
$lblMicroExec.Text = "Testes de Latência: 0/100"
$lblMicroExec.Font = New-Object System.Drawing.Font("Segoe UI", 11)
$lblMicroExec.Location = New-Object System.Drawing.Point(50, 210)
$lblMicroExec.AutoSize = $true

$pbMicro = New-Object System.Windows.Forms.ProgressBar
$pbMicro.Location = New-Object System.Drawing.Point(50, 240)
$pbMicro.Maximum = 100
$pbMicro.Height = 20

# --- AVISO DE SERVIDORES SEM RESPOSTA ---
$lblAviso100 = New-Object System.Windows.Forms.Label
$lblAviso100.Text = "Nota: É normal que alguns testes parem em '1/100'. Isso apenas indica um servidor sem resposta, que será pulado automaticamente."
$lblAviso100.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblAviso100.ForeColor = [System.Drawing.Color]::DimGray
$lblAviso100.Location = New-Object System.Drawing.Point(50, 275)
$lblAviso100.AutoSize = $true

$btnCancelExec = New-Object System.Windows.Forms.Button
$btnCancelExec.Text = "Cancelar Teste"
$btnCancelExec.Size = New-Object System.Drawing.Size(180, 45)
$btnCancelExec.BackColor = [System.Drawing.Color]::IndianRed
$btnCancelExec.ForeColor = [System.Drawing.Color]::White
$btnCancelExec.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnCancelExec.FlatAppearance.BorderSize = 0
$btnCancelExec.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
$btnCancelExec.Cursor = [System.Windows.Forms.Cursors]::Hand

$pnlExec.Add_Resize({
    $btnCancelExec.Left = ($pnlExec.Width - $btnCancelExec.Width) / 2
    $btnCancelExec.Top = 350
    $pbMacro.Width = $pnlExec.Width - 100 
    $pbMicro.Width = $pnlExec.Width - 100
})

$pnlExec.Controls.Add($lblMacroExec)
$pnlExec.Controls.Add($pbMacro)
$pnlExec.Controls.Add($lblMicroExec)
$pnlExec.Controls.Add($pbMicro)
$pnlExec.Controls.Add($lblAviso100)
$pnlExec.Controls.Add($btnCancelExec)
$form.Controls.Add($pnlExec)

# --- PAINEL DE RESULTADOS ---
$pnlRes = New-Object System.Windows.Forms.Panel
$pnlRes.Dock = "Fill"
$pnlRes.Visible = $false

$gridRes = New-Object System.Windows.Forms.DataGridView
$gridRes.Dock = "Fill"
$gridRes.AutoSizeColumnsMode = "Fill"
$gridRes.AllowUserToAddRows = $false
$gridRes.ReadOnly = $true
$gridRes.AllowUserToResizeRows = $false
$gridRes.RowHeadersVisible = $false
$gridRes.BackgroundColor = "White"
$gridRes.BorderStyle = "None"
$gridRes.GridColor = [System.Drawing.Color]::LightGray
$gridRes.DefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$gridRes.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$gridRes.ColumnHeadersBorderStyle = "Single"

$gridRes.SelectionMode = "FullRowSelect"
$gridRes.MultiSelect = $false
$gridRes.Add_SelectionChanged({ 
    if ($gridRes.SelectedCells.Count -gt 0) { $gridRes.ClearSelection() } 
})

$gridRes.Add_CellFormatting({
    param($sender, $e)
    if ($e.RowIndex -ge 0 -and $e.RowIndex -lt $sender.Rows.Count) {
        $row = $sender.Rows[$e.RowIndex]
        $mediaStr = $row.Cells["Média"].Value
        
        if ($mediaStr -eq "--------") {
            $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::Salmon
        } else {
            $media = $row.Cells["RawMedia"].Value
            $medMax = $row.Cells["RawMedMax"].Value
            $max = $row.Cells["RawMax"].Value

            if ($media -le 60 -and $max -lt 130 -and ($medMax -eq 0 -or $medMax -lt 80)) {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::LightGreen
            } elseif ($media -ge 80 -or $medMax -ge 80 -or $max -ge 130) {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::LightGray
            } else {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::White
            }
        }
        $row.DefaultCellStyle.SelectionBackColor = $row.DefaultCellStyle.BackColor
        $row.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::Black
    }
})

# ==============================================================================
# GERADOR DO GRÁFICO DE ESTABILIDADE
# ==============================================================================
function Show-LatencyChart {
    param([string]$Nome, [string]$IP, [array]$Tempos)
    
    $media = ($Tempos | Measure-Object -Average).Average
    
    # --------------------------------------------------------------------------
    # GESTÃO DE FONTES (Prevenção de Memory Leak)
    # --------------------------------------------------------------------------
    $fntNormal = New-Object System.Drawing.Font("Segoe UI", 10)
    $fntBold   = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    
    $frmChart = New-Object System.Windows.Forms.Form
    $frmChart.Text = "Análise de Estabilidade - $Nome ($IP)"
    $frmChart.StartPosition = "CenterParent"
    $frmChart.BackColor = "White"
    
    # --------------------------------------------------------------------------
    # COMPATIBILIDADE COM ESCALA DE TELA (DPI)
    # --------------------------------------------------------------------------
    $frmChart.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::None

    $chart = New-Object System.Windows.Forms.DataVisualization.Charting.Chart
    $chart.BackColor = "White"

    $legend = New-Object System.Windows.Forms.DataVisualization.Charting.Legend
    $legend.Name = "Legend1"
    $legend.Docking = [System.Windows.Forms.DataVisualization.Charting.Docking]::Bottom
    $legend.Alignment = [System.Drawing.StringAlignment]::Center
    $legend.Font = $fntNormal
    $chart.Legends.Add($legend)

    $area = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea
    $area.Name = "MainArea"
    $area.AxisX.Title = "Número do Teste (1 a 100)"
    $area.AxisX.TitleFont = $fntBold 
    
    # --- MATEMÁTICA DO EIXO X LIMPA E FIXA ---
    # Isso garante os marcadores exatos em: 0, 20, 40, 60, 80 e 100.
    $area.AxisX.Minimum = 0
    $area.AxisX.Maximum = $Tempos.Count
    $area.AxisX.Interval = 20

    $area.AxisY.Title = "Tempo de Resposta (ms)"
    $area.AxisY.TitleFont = $fntBold
    $area.AxisX.MajorGrid.LineColor = [System.Drawing.Color]::LightGray
    $area.AxisY.MajorGrid.LineColor = [System.Drawing.Color]::LightGray
    $chart.ChartAreas.Add($area)

    $linhaMedia = New-Object System.Windows.Forms.DataVisualization.Charting.StripLine
    $linhaMedia.IntervalOffset = $media 
    $linhaMedia.StripWidth = 0 
    $linhaMedia.BorderColor = [System.Drawing.Color]::DarkOrange 
    $linhaMedia.BorderDashStyle = [System.Windows.Forms.DataVisualization.Charting.ChartDashStyle]::Dash 
    $linhaMedia.BorderWidth = 2
    $area.AxisY.StripLines.Add($linhaMedia)

    $serie = New-Object System.Windows.Forms.DataVisualization.Charting.Series
    $serie.Name = "Tempo Atual de Resposta" 
    $serie.ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Line
    $serie.Color = [System.Drawing.Color]::SteelBlue
    $serie.BorderWidth = 2
    $serie.IsValueShownAsLabel = $false
    $chart.Series.Add($serie)

    $serieMediaLegenda = New-Object System.Windows.Forms.DataVisualization.Charting.Series
    $serieMediaLegenda.Name = "Tempo Médio"
    $serieMediaLegenda.ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Line
    $serieMediaLegenda.Color = [System.Drawing.Color]::DarkOrange
    $serieMediaLegenda.BorderWidth = 2
    $serieMediaLegenda.BorderDashStyle = [System.Windows.Forms.DataVisualization.Charting.ChartDashStyle]::Dash
    $chart.Series.Add($serieMediaLegenda)

    # ==========================================================================
    # A MATEMÁTICA PERFEITA DE DISTRIBUIÇÃO (Pixel-Perfect Edge)
    # O primeiro ponto é deslocado 0.3 unidades para a direita. 
    # Isso garante que a linha encoste APENAS no LADO DIREITO do marcador 0, 
    # sem invadir ou pintar o meio da linha de grade do eixo Y.
    # Todos os outros pontos são esticados milimetricamente para tocarem no 100.
    # ==========================================================================
    if ($Tempos.Count -gt 1) {
        $startX = 0.3
        $endX = $Tempos.Count
        $stepX = ($endX - $startX) / ($Tempos.Count - 1)
        
        for ($i = 0; $i -lt $Tempos.Count; $i++) {
            $currentX = $startX + ($i * $stepX)
            $serie.Points.AddXY($currentX, $Tempos[$i]) | Out-Null
        }
    }

    $lblMedia = New-Object System.Windows.Forms.Label
    $lblMedia.Text = "$([math]::Round($media, 2)) ms"
    $lblMedia.ForeColor = [System.Drawing.Color]::DarkOrange
    $lblMedia.Font = $fntBold
    $lblMedia.AutoSize = $true

    $larguraTexto = $lblMedia.PreferredSize.Width
    $larguraGrafico = 800
    
    $frmChart.ClientSize = New-Object System.Drawing.Size(($larguraGrafico + $larguraTexto + 20), 450)

    $chart.Dock = "None"
    $chart.Location = New-Object System.Drawing.Point(0, 0)
    $chart.Size = New-Object System.Drawing.Size($larguraGrafico, 450)
    $chart.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right

    $frmChart.Controls.Add($lblMedia)
    $frmChart.Controls.Add($chart)

    $atualizarPosicao = {
        if ($frmChart.WindowState -eq 'Minimized') { return }
        $chart.Update()
        try {
            $yPixel = $chart.ChartAreas["MainArea"].AxisY.ValueToPixelPosition($media)
            $xPixel = $chart.ChartAreas["MainArea"].AxisX.ValueToPixelPosition($Tempos.Count)
            
            $lblMedia.Top = [int]$yPixel - ($lblMedia.Height / 2)
            $lblMedia.Left = [int]$xPixel + 4
        } catch { }
    }

    $frmChart.Add_Shown($atualizarPosicao)
    $frmChart.Add_Resize($atualizarPosicao)

    $frmChart.ShowDialog() | Out-Null
    $frmChart.Dispose()
    
    # --------------------------------------------------------------------------
    # Liberação dos recursos gráficos da memória
    # --------------------------------------------------------------------------
    $fntNormal.Dispose()
    $fntBold.Dispose()
}

# ==============================================================================
# INTERAÇÃO DA TABELA (CLIQUE ESQUERDO / DIREITO)
# ==============================================================================
$gridRes.Add_CellMouseClick({
    param($sender, $e)
    
    if ($e.RowIndex -ge 0) {
        $row = $sender.Rows[$e.RowIndex]
        $ipClicado = $row.Cells["IP"].Value
        
        # --- CLIQUE ESQUERDO: COPIAR IP ---
        if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
            if (-not [string]::IsNullOrWhiteSpace($ipClicado)) {
                [System.Windows.Forms.Clipboard]::SetText($ipClicado)
                [void][System.Windows.Forms.MessageBox]::Show("O IP '$ipClicado' foi copiado para a área de transferência com sucesso!", "IP Copiado", 0, "Information")
            }
        }
        # --- CLIQUE DIREITO: ABRIR GRÁFICO ---
        elseif ($e.Button -eq [System.Windows.Forms.MouseButtons]::Right) {
            $sender.ClearSelection()
            $row.Selected = $true # Move a seleção visual para a linha clicada com o botão direito
            
            $mediaStr = $row.Cells["Média"].Value
            $temposBrutos = $row.Cells["TemposBrutos"].Value
            $nomeClicado = $row.Cells["Nome"].Value
            
            # Trava de segurança: Se o servidor falhou ou não tem 100 testes, não desenha o gráfico.
            if ($mediaStr -eq "--------" -or $null -eq $temposBrutos -or $temposBrutos.Count -eq 0) {
                [void][System.Windows.Forms.MessageBox]::Show("Não há dados de estabilidade suficientes para gerar um gráfico deste servidor, pois ele falhou no teste.", "Gráfico Indisponível", 0, "Warning")
                return
            }
            
            Show-LatencyChart -Nome $nomeClicado -IP $ipClicado -Tempos $temposBrutos
        }
    }
})

# ==============================================================================
# DICA CUSTOMIZADA (TOOLTIP) ATUALIZADA
# ==============================================================================
$gridRes.Add_CellToolTipTextNeeded({
    param($sender, $e)
    if ($e.RowIndex -ge 0) {
        $e.ToolTipText = "• Clique ESQUERDO para copiar o endereço de IP.`n• Clique DIREITO para abrir o Gráfico de Estabilidade."
    }
})

# ==============================================================================
# AJUSTE AUTOMÁTICO DAS COLUNAS DE TEMPO
# ==============================================================================
$gridRes.Add_DataBindingComplete({
    param($sender, $e)
    
    # Lista com o nome exato das colunas que queremos espremer
    $colunasTempo = @("Mínimo", "Média-Mínima", "Média", "Média-Máxima", "Máximo")
    
    foreach ($colName in $colunasTempo) {
        if ($sender.Columns.Contains($colName)) {
            # Ajusta a largura baseada no maior tamanho (seja o texto do cabeçalho ou o valor em ms)
            $sender.Columns[$colName].AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::AllCells
        }
    }
})

$pnlResBot = New-Object System.Windows.Forms.Panel
$pnlResBot.Dock = "Bottom"
$pnlResBot.Height = 150
$pnlResBot.BackColor = "White"

$txtResInfos = New-Object System.Windows.Forms.RichTextBox
$txtResInfos.Dock = "Fill"
$txtResInfos.ReadOnly = $true
$txtResInfos.BackColor = "White"
$txtResInfos.Font = New-Object System.Drawing.Font("Consolas", 10) # Mantido Consolas apenas para o alinhamento perfeito do ranking
$txtResInfos.BorderStyle = "None"

$pnlResBotBtns = New-Object System.Windows.Forms.Panel
$pnlResBotBtns.Dock = "Right"
$pnlResBotBtns.Width = 220
$pnlResBotBtns.BackColor = "White"

# ==============================================================================
# BOTÃO DE FILTROS E MENU SUSPENSO
# ==============================================================================
$btnFiltroRes = New-Object System.Windows.Forms.Button
$btnFiltroRes.Text = "Filtros ▾"
$btnFiltroRes.Location = New-Object System.Drawing.Point(30, 15)
$btnFiltroRes.Size = New-Object System.Drawing.Size(160, 35)
$btnFiltroRes.BackColor = [System.Drawing.Color]::MediumPurple
$btnFiltroRes.ForeColor = [System.Drawing.Color]::White
$btnFiltroRes.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnFiltroRes.FlatAppearance.BorderSize = 0
$btnFiltroRes.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$btnFiltroRes.Cursor = [System.Windows.Forms.Cursors]::Hand

$ctxFiltro = New-Object System.Windows.Forms.ContextMenuStrip
$btnFiltroRes.Add_Click({ $ctxFiltro.Show($btnFiltroRes, (New-Object System.Drawing.Point(0, $btnFiltroRes.Height))) })

# ==============================================================================
# BOTÕES ANTIGOS (REDIMENSIONADOS E REPOSICIONADOS)
# ==============================================================================
$btnSaveRes = New-Object System.Windows.Forms.Button
$btnSaveRes.Text = "Salvar Arquivo"
$btnSaveRes.Location = New-Object System.Drawing.Point(30, 60) # Movido para baixo
$btnSaveRes.Size = New-Object System.Drawing.Size(160, 35) # Altura reduzida para 35
$btnSaveRes.BackColor = [System.Drawing.Color]::SteelBlue
$btnSaveRes.ForeColor = [System.Drawing.Color]::White
$btnSaveRes.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnSaveRes.FlatAppearance.BorderSize = 0
$btnSaveRes.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$btnSaveRes.Cursor = [System.Windows.Forms.Cursors]::Hand

$btnVoltarRes = New-Object System.Windows.Forms.Button
$btnVoltarRes.Text = "Voltar"
$btnVoltarRes.Location = New-Object System.Drawing.Point(30, 105) # Movido mais para baixo
$btnVoltarRes.Size = New-Object System.Drawing.Size(160, 35) # Altura reduzida para 35
$btnVoltarRes.BackColor = [System.Drawing.Color]::SlateGray
$btnVoltarRes.ForeColor = [System.Drawing.Color]::White
$btnVoltarRes.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnVoltarRes.FlatAppearance.BorderSize = 0
$btnVoltarRes.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$btnVoltarRes.Cursor = [System.Windows.Forms.Cursors]::Hand

# Adiciona os 3 botões ao painel lateral
$pnlResBotBtns.Controls.Add($btnFiltroRes)
$pnlResBotBtns.Controls.Add($btnSaveRes)
$pnlResBotBtns.Controls.Add($btnVoltarRes)

$pnlResBot.Controls.Add($txtResInfos)
$pnlResBot.Controls.Add($pnlResBotBtns)

$pnlRes.Controls.Add($pnlResBot)
$pnlRes.Controls.Add($gridRes)
$gridRes.BringToFront()
$form.Controls.Add($pnlRes)

# ==============================================================================
# VARIÁVEIS DO ESCUDO DE SELEÇÃO
# ==============================================================================
$global:MouseSobreCheckbox = $false
$global:RevertendoSelecao = $false
$global:LinhaValidaIPv4 = -1
$global:LinhaValidaIPv6 = -1

# ==============================================================================
# 5. DATA GRIDS DO MENU PRINCIPAL E LISTA NEGRA
# ==============================================================================
function Apply-GridSecurity ($grid) {
    # --- ACELERAÇÃO GRÁFICA ---
    $type = $grid.GetType()
    $prop = $type.GetProperty("DoubleBuffered", [System.Reflection.BindingFlags]::Instance -bor [System.Reflection.BindingFlags]::NonPublic)
    $prop.SetValue($grid, $true, $null)

    $grid.Dock = "Fill"; $grid.AutoSizeColumnsMode = "Fill"; $grid.AllowUserToAddRows = $false
    $grid.SelectionMode = "FullRowSelect"; 
    $grid.ReadOnly = $true
    $grid.MultiSelect = $false
    $grid.AllowUserToResizeColumns = $false; $grid.AllowUserToResizeRows = $false
    $grid.ColumnHeadersHeightSizeMode = "DisableResizing"; $grid.RowHeadersVisible = $false
    $grid.BackgroundColor = [System.Drawing.Color]::White; $grid.EnableHeadersVisualStyles = $false
    $grid.BorderStyle = "None"
    $grid.GridColor = [System.Drawing.Color]::LightGray
    $grid.ShowCellToolTips = $false

    $grid.DefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $grid.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $grid.ColumnHeadersBorderStyle = "Single"
    $grid.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::WhiteSmoke
    $grid.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::SteelBlue
    $grid.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::White

    # 1. Detecta mouse sobre o checkbox
    $grid.Add_CellMouseEnter({
        param($sender, $e)
        if ($e.ColumnIndex -ge 0 -and $sender.Columns[$e.ColumnIndex].Name -eq "Testar") {
            $global:MouseSobreCheckbox = $true
        }
    })
    $grid.Add_CellMouseLeave({
        param($sender, $e)
        if ($e.ColumnIndex -ge 0 -and $sender.Columns[$e.ColumnIndex].Name -eq "Testar") {
            $global:MouseSobreCheckbox = $false
        }
    })

    # 2. Previne seleção acidental da linha inteira
    $grid.Add_SelectionChanged({
        param($sender, $e)
        if ($global:RevertendoSelecao) { return }

        if ($global:MouseSobreCheckbox) {
            $global:RevertendoSelecao = $true
            $linhaValida = if ($sender -eq $gridIPv4) { $global:LinhaValidaIPv4 } else { $global:LinhaValidaIPv6 }

            if ($linhaValida -ge 0 -and $linhaValida -lt $sender.Rows.Count) {
                # Previne o flicker visual da tabela:
                # Ao invés de mudar o CurrentCell (que forçava o auto-scroll para cima),
                # nós apenas mantemos a propriedade "Selected" acesa na linha original.
                $sender.ClearSelection()
                $sender.Rows[$linhaValida].Selected = $true
            } else {
                $sender.ClearSelection()
            }
            $global:RevertendoSelecao = $false
        } else {
            # Se clicou no Nome ou IP, a seleção é verdadeira e será memorizada
            if ($sender.CurrentCell -ne $null) {
                if ($sender -eq $gridIPv4) { $global:LinhaValidaIPv4 = $sender.CurrentCell.RowIndex }
                else { $global:LinhaValidaIPv6 = $sender.CurrentCell.RowIndex }
            }
        }
    })

    # 3. Transfere a ação de clique para CellMouseUp
    $grid.Add_CellMouseUp({
        param($sender, $e)
        if ($e.Button -eq 'Left' -and $e.RowIndex -ge 0 -and $sender.Columns[$e.ColumnIndex].Name -eq "Testar") {
            $row = $sender.Rows[$e.RowIndex]
            $item = $row.DataBoundItem
            
            $priRuim = ($null -ne $item.DNS_Primario -and $global:BlacklistAtiva -contains $item.DNS_Primario)
            $secRuim = ($null -ne $item.DNS_Secundario -and $global:BlacklistAtiva -contains $item.DNS_Secundario)
            
            $temPriValido = (-not [string]::IsNullOrWhiteSpace($item.DNS_Primario) -and -not $priRuim)
            $temSecValido = (-not [string]::IsNullOrWhiteSpace($item.DNS_Secundario) -and -not $secRuim)
            
            # Só alterna o valor se o DNS for válido
            if ($temPriValido -or $temSecValido) {
                if ($item.Selecionado -eq "True") { $item.Selecionado = "False" } else { $item.Selecionado = "True" }
                $sender.InvalidateRow($e.RowIndex) 
            }
            
            if ($sender -eq $gridIPv4) { Salvar-Checkboxes $true } else { Salvar-Checkboxes $false }
        }
    })
}

# 1. PRIMEIRO: Criamos as tabelas e aplicamos a segurança
$gridIPv4 = New-Object System.Windows.Forms.DataGridView; Apply-GridSecurity $gridIPv4
$gridIPv6 = New-Object System.Windows.Forms.DataGridView; Apply-GridSecurity $gridIPv6

# 2. SEGUNDO: Agora que elas existem, amarramos o motor de pintura nelas!
$gridIPv4.Add_CellFormatting($script:ColorirListaNegra)
$gridIPv6.Add_CellFormatting($script:ColorirListaNegra)

$global:BlacklistAtiva = @()

function Load-CsvToGrid {
    $dataIPv4 = Import-Csv -Path $fileIPv4 -Encoding UTF8 | Sort-Object Nome
    $dataIPv6 = Import-Csv -Path $fileIPv6 -Encoding UTF8 | Sort-Object Nome
    
    $global:BlacklistAtiva = @()
    if (Test-Path $fileBlacklist) {
        $blData = @(Import-Csv -Path $fileBlacklist -Encoding UTF8)
        $global:BlacklistAtiva = $blData | Where-Object { $_.Status -eq "Ativo" } | Select-Object -ExpandProperty IP
    }

    $gridIPv4.DataSource = [System.Collections.ArrayList]@($dataIPv4)
    $gridIPv6.DataSource = [System.Collections.ArrayList]@($dataIPv6)
    
    foreach ($grid in @($gridIPv4, $gridIPv6)) {
        if (-not $grid.Columns.Contains("Testar")) {
            $colCheck = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
            $colCheck.Name = "Testar"; $colCheck.HeaderText = "☑"
            
            # Ajuste dinâmico de largura da coluna
            $colCheck.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::AllCells 
            
            # Centraliza o conteúdo das linhas
            $colCheck.DefaultCellStyle.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleCenter
            
            # Centraliza o texto do cabeçalho
            $colCheck.HeaderCell.Style.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleCenter
            
            # Tamanho 14 para o checkbox unicode ficar perfeito
            $colCheck.DefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 14) 
            $grid.Columns.Insert(0, $colCheck)
        }

        if ($grid.Columns.Contains("DNS_Primario")) { $grid.Columns["DNS_Primario"].HeaderText = "DNS Primário" }
        if ($grid.Columns.Contains("DNS_Secundario")) { $grid.Columns["DNS_Secundario"].HeaderText = "DNS Secundário" }
        if ($grid.Columns.Contains("Selecionado")) { $grid.Columns["Selecionado"].Visible = $false }

        foreach ($col in $grid.Columns) { $col.SortMode = "NotSortable" }
    }
}

# --- O MOTOR DE PINTURA FINAL ---
$script:ColorirListaNegra = {
    param($sender, $e)
    if ($e.RowIndex -ge 0 -and $e.RowIndex -lt $sender.Rows.Count) {
        $item = $sender.Rows[$e.RowIndex].DataBoundItem
        if ($null -eq $item) { return }

        $pri = $item.DNS_Primario
        $sec = $item.DNS_Secundario
        
        $priRuim = ($null -ne $pri -and $global:BlacklistAtiva -contains $pri)
        $secRuim = ($null -ne $sec -and $global:BlacklistAtiva -contains $sec)
        
        $temPriValido = (-not [string]::IsNullOrWhiteSpace($pri) -and -not $priRuim)
        $temSecValido = (-not [string]::IsNullOrWhiteSpace($sec) -and -not $secRuim)
        
        $totalmenteMorto = (-not $temPriValido -and -not $temSecValido)
        $nomeColuna = $sender.Columns[$e.ColumnIndex].Name
        
        $corFundoNormal = [System.Drawing.Color]::White
        $corFundoSelecao = [System.Drawing.Color]::SteelBlue
        $corTextoSelecao = [System.Drawing.Color]::White
        
        $celulaComProblema = $false
        
        if ($totalmenteMorto) {
            $celulaComProblema = $true
        } else {
            if ($nomeColuna -eq "DNS_Primario" -and $priRuim) { $celulaComProblema = $true }
            if ($nomeColuna -eq "DNS_Secundario" -and $secRuim) { $celulaComProblema = $true }
        }
        
        if ($celulaComProblema) {
            $corFundoNormal = [System.Drawing.Color]::MistyRose
            $corFundoSelecao = [System.Drawing.Color]::Indigo
        }
        
        if ($nomeColuna -eq "Testar") {
            if ($totalmenteMorto) {
                $e.Value = "☒"
                $e.CellStyle.ForeColor = [System.Drawing.Color]::IndianRed
                $corTextoSelecao = [System.Drawing.Color]::White 
            } else {
                if ($item.Selecionado -eq "True") {
                    $e.Value = "☑"
                    $e.CellStyle.ForeColor = [System.Drawing.Color]::MediumSeaGreen
                } else {
                    $e.Value = "☐"
                    $e.CellStyle.ForeColor = [System.Drawing.Color]::Gray
                }
            }
            $e.FormattingApplied = $true
        }
        
        $e.CellStyle.BackColor = $corFundoNormal
        $e.CellStyle.SelectionBackColor = $corFundoSelecao
        $e.CellStyle.SelectionForeColor = $corTextoSelecao
    }
}

$gridIPv4.Add_CellFormatting($script:ColorirListaNegra)
$gridIPv6.Add_CellFormatting($script:ColorirListaNegra)

function Salvar-Checkboxes {
    param($isIPv4)
    $grid = if ($isIPv4) { $gridIPv4 } else { $gridIPv6 }
    $file = if ($isIPv4) { $fileIPv4 } else { $fileIPv6 }
    
    $dadosParaSalvar = New-Object System.Collections.ArrayList
    foreach ($row in $grid.Rows) {
        if ($row.Index -ge 0) {
            $item = $row.DataBoundItem
            $temPri = (-not [string]::IsNullOrWhiteSpace($item.DNS_Primario) -and $global:BlacklistAtiva -notcontains $item.DNS_Primario)
            $temSec = (-not [string]::IsNullOrWhiteSpace($item.DNS_Secundario) -and $global:BlacklistAtiva -notcontains $item.DNS_Secundario)
            if (-not $temPri -and -not $temSec) { $item.Selecionado = "False" }
            
            $dadosParaSalvar.Add($item) | Out-Null
        }
    }
    $dadosParaSalvar | Export-Csv -Path $file -NoTypeInformation -Encoding UTF8
}

$pnlBotIPv4 = New-Object System.Windows.Forms.Panel; $pnlBotIPv4.Dock = "Bottom"; $pnlBotIPv4.Height = 70; $pnlBotIPv4.BackColor = "White"
$pnlBotIPv6 = New-Object System.Windows.Forms.Panel; $pnlBotIPv6.Dock = "Bottom"; $pnlBotIPv6.Height = 70; $pnlBotIPv6.BackColor = "White"

function New-FlatButton ($text, $x, $y, $w, $bgCor) {
    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = $text; $btn.Location = New-Object System.Drawing.Point($x, $y); $btn.Size = New-Object System.Drawing.Size($w, 40)
    $btn.FlatStyle = "Flat"; $btn.FlatAppearance.BorderSize = 0; $btn.BackColor = $bgCor; $btn.ForeColor = "White"
    $btn.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold); $btn.Cursor = [System.Windows.Forms.Cursors]::Hand
    return $btn
}

# ==============================================================================
# BOTÕES DOS PAINÉIS (COM NOMES ATUALIZADOS E COORDENADAS RECALCULADAS)
# ==============================================================================
$btnAddIPv4 = New-FlatButton "Adicionar DNS IPv4" 15 15 140 "SteelBlue"
$btnEditIPv4 = New-FlatButton "Editar DNS" 165 15 100 "SlateGray"
$btnRemIPv4 = New-FlatButton "Remover DNS" 275 15 110 "IndianRed"
$btnBlacklistIPv4 = New-FlatButton "Lista Negra" 395 15 100 "Purple"
$btnTestIPv4 = New-FlatButton "INICIAR TESTE" 505 15 130 "MediumSeaGreen"

$pnlBotIPv4.Controls.Add($btnAddIPv4); $pnlBotIPv4.Controls.Add($btnEditIPv4); $pnlBotIPv4.Controls.Add($btnRemIPv4); $pnlBotIPv4.Controls.Add($btnBlacklistIPv4); $pnlBotIPv4.Controls.Add($btnTestIPv4)
$pnlContentIPv4.Controls.Add($pnlBotIPv4); $pnlContentIPv4.Controls.Add($gridIPv4); $gridIPv4.BringToFront()

$btnAddIPv6 = New-FlatButton "Adicionar DNS IPv6" 15 15 140 "SteelBlue"
$btnEditIPv6 = New-FlatButton "Editar DNS" 165 15 100 "SlateGray"
$btnRemIPv6 = New-FlatButton "Remover DNS" 275 15 110 "IndianRed"
$btnBlacklistIPv6 = New-FlatButton "Lista Negra" 395 15 100 "Purple"
$btnTestIPv6 = New-FlatButton "INICIAR TESTE" 505 15 130 "MediumSeaGreen"

$pnlBotIPv6.Controls.Add($btnAddIPv6); $pnlBotIPv6.Controls.Add($btnEditIPv6); $pnlBotIPv6.Controls.Add($btnRemIPv6); $pnlBotIPv6.Controls.Add($btnBlacklistIPv6); $pnlBotIPv6.Controls.Add($btnTestIPv6)
$pnlContentIPv6.Controls.Add($pnlBotIPv6); $pnlContentIPv6.Controls.Add($gridIPv6); $gridIPv6.BringToFront()

function Set-CheckboxesTopo ($estado) {
    $gridAtivo = if ($pnlContentIPv4.Visible) { $gridIPv4 } else { $gridIPv6 }
    foreach ($row in $gridAtivo.Rows) {
        if ($row.Index -ge 0) { 
            $item = $row.DataBoundItem
            $temPri = (-not [string]::IsNullOrWhiteSpace($item.DNS_Primario) -and $global:BlacklistAtiva -notcontains $item.DNS_Primario)
            $temSec = (-not [string]::IsNullOrWhiteSpace($item.DNS_Secundario) -and $global:BlacklistAtiva -notcontains $item.DNS_Secundario)
            
            if ($temPri -or $temSec) {
                if ($estado) { $item.Selecionado = "True" } else { $item.Selecionado = "False" }
            } else {
                $item.Selecionado = "False"
            }
        }
    }
    $gridAtivo.Invalidate() 
    Salvar-Checkboxes $pnlContentIPv4.Visible 
}

$btnMarcarTopo.Add_Click({ Set-CheckboxesTopo $true })
$btnDesmarcarTopo.Add_Click({ Set-CheckboxesTopo $false })

# ==============================================================================
# MOTORES DE SINCRONIZAÇÃO DA LISTA NEGRA (EFEITO CASCATA)
# ==============================================================================
function Sync-RemoveBlacklist ($ip1, $ip2) {
    if (-not (Test-Path $fileBlacklist)) { return }
    $bl = @(Import-Csv $fileBlacklist -Encoding UTF8)
    $bl = $bl | Where-Object { $_.IP -ne $ip1 -and $_.IP -ne $ip2 }
    if ($bl.Count -gt 0) { $bl | Export-Csv $fileBlacklist -NoTypeInformation -Encoding UTF8 } 
    else { "IP,Nome,Status" | Out-File $fileBlacklist -Encoding UTF8 }
}

function Sync-EditBlacklist ($oldPri, $oldSec, $newPri, $newSec, $newName) {
    if (-not (Test-Path $fileBlacklist)) { return }
    $bl = @(Import-Csv $fileBlacklist -Encoding UTF8)
    $changed = $false; $novaLista = New-Object System.Collections.ArrayList
    
    foreach ($item in $bl) {
        if ($item.IP -eq $oldPri -and -not [string]::IsNullOrWhiteSpace($oldPri)) {
            if (-not [string]::IsNullOrWhiteSpace($newPri)) {
                $item.IP = $newPri; $item.Nome = $newName
                $novaLista.Add($item) | Out-Null
            }
            $changed = $true
        } elseif ($item.IP -eq $oldSec -and -not [string]::IsNullOrWhiteSpace($oldSec)) {
            if (-not [string]::IsNullOrWhiteSpace($newSec)) {
                $item.IP = $newSec; $item.Nome = $newName
                $novaLista.Add($item) | Out-Null
            }
            $changed = $true
        } else {
            $novaLista.Add($item) | Out-Null
        }
    }
    
    if ($changed) {
        if ($novaLista.Count -gt 0) { $novaLista | Export-Csv $fileBlacklist -NoTypeInformation -Encoding UTF8 } 
        else { "IP,Nome,Status" | Out-File $fileBlacklist -Encoding UTF8 }
    }
}

# ==============================================================================
# EVENTOS DOS BOTÕES (COM SINCRONIZAÇÃO E BLINDAGEM DE IP E REPETIÇÕES)
# ==============================================================================
$btnAddIPv4.Add_Click({ 
    $n = Show-DNSDialog -Titulo "Novo IPv4" -ListaExistente @($gridIPv4.DataSource) -TipoIP "IPv4"
    if ($n) { 
        $n | Add-Member -MemberType NoteProperty -Name "Selecionado" -Value "True"
        $dados = @(Import-Csv $fileIPv4 -Encoding UTF8); $dados += $n
        $dados | Sort-Object Nome | Export-Csv $fileIPv4 -NoTypeInformation -Encoding UTF8; Load-CsvToGrid 
    } 
})

$btnAddIPv6.Add_Click({ 
    $n = Show-DNSDialog -Titulo "Novo IPv6" -ListaExistente @($gridIPv6.DataSource) -TipoIP "IPv6"
    if ($n) { 
        $n | Add-Member -MemberType NoteProperty -Name "Selecionado" -Value "True"
        $dados = @(Import-Csv $fileIPv6 -Encoding UTF8); $dados += $n
        $dados | Sort-Object Nome | Export-Csv $fileIPv6 -NoTypeInformation -Encoding UTF8; Load-CsvToGrid 
    } 
})

$btnEditIPv4.Add_Click({ 
    if ($gridIPv4.SelectedRows.Count -gt 0) { 
        $r = $gridIPv4.SelectedRows[0].DataBoundItem; 
        $oldPri = $r.DNS_Primario; $oldSec = $r.DNS_Secundario;
        
        $e = Show-DNSDialog -Titulo "Editar IPv4" -NomeAtual $r.Nome -PrimarioAtual $r.DNS_Primario -SecundarioAtual $r.DNS_Secundario -ListaExistente @($gridIPv4.DataSource) -TipoIP "IPv4"
        
        if ($e) { 
            Sync-EditBlacklist $oldPri $oldSec $e.DNS_Primario $e.DNS_Secundario $e.Nome
            $dados = @(Import-Csv $fileIPv4 -Encoding UTF8)
            foreach ($i in $dados) { if ($i.Nome -eq $r.Nome) { $i.Nome = $e.Nome; $i.DNS_Primario = $e.DNS_Primario; $i.DNS_Secundario = $e.DNS_Secundario } }
            $dados | Sort-Object Nome | Export-Csv $fileIPv4 -NoTypeInformation -Encoding UTF8; Load-CsvToGrid 
        } 
    } 
})

$btnEditIPv6.Add_Click({ 
    if ($gridIPv6.SelectedRows.Count -gt 0) { 
        $r = $gridIPv6.SelectedRows[0].DataBoundItem; 
        $oldPri = $r.DNS_Primario; $oldSec = $r.DNS_Secundario;
        
        $e = Show-DNSDialog -Titulo "Editar IPv6" -NomeAtual $r.Nome -PrimarioAtual $r.DNS_Primario -SecundarioAtual $r.DNS_Secundario -ListaExistente @($gridIPv6.DataSource) -TipoIP "IPv6"
        
        if ($e) { 
            Sync-EditBlacklist $oldPri $oldSec $e.DNS_Primario $e.DNS_Secundario $e.Nome
            $dados = @(Import-Csv $fileIPv6 -Encoding UTF8)
            foreach ($i in $dados) { if ($i.Nome -eq $r.Nome) { $i.Nome = $e.Nome; $i.DNS_Primario = $e.DNS_Primario; $i.DNS_Secundario = $e.DNS_Secundario } }
            $dados | Sort-Object Nome | Export-Csv $fileIPv6 -NoTypeInformation -Encoding UTF8; Load-CsvToGrid 
        } 
    } 
})

$btnRemIPv4.Add_Click({ 
    if ($gridIPv4.SelectedRows.Count -gt 0) { 
        $r = $gridIPv4.SelectedRows[0].DataBoundItem; 
        $msg = [System.Windows.Forms.MessageBox]::Show("Deseja realmente remover o DNS '$($r.Nome)' da lista?", "Confirmar Exclusão", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning); 
        if ($msg -eq "Yes") { 
            Sync-RemoveBlacklist $r.DNS_Primario $r.DNS_Secundario
            $dados = @(Import-Csv $fileIPv4 -Encoding UTF8); $dadosRestantes = @($dados | Where-Object { $_.Nome -ne $r.Nome }); 
            if ($dadosRestantes.Count -gt 0) { $dadosRestantes | Sort-Object Nome | Export-Csv $fileIPv4 -NoTypeInformation -Encoding UTF8 } else { "Nome,DNS_Primario,DNS_Secundario,Selecionado" | Out-File $fileIPv4 -Encoding UTF8 }; 
            Load-CsvToGrid 
        } 
    } 
})

$btnRemIPv6.Add_Click({ 
    if ($gridIPv6.SelectedRows.Count -gt 0) { 
        $r = $gridIPv6.SelectedRows[0].DataBoundItem; 
        $msg = [System.Windows.Forms.MessageBox]::Show("Deseja realmente remover o DNS '$($r.Nome)' da lista?", "Confirmar Exclusão", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning); 
        if ($msg -eq "Yes") { 
            Sync-RemoveBlacklist $r.DNS_Primario $r.DNS_Secundario
            $dados = @(Import-Csv $fileIPv6 -Encoding UTF8); $dadosRestantes = @($dados | Where-Object { $_.Nome -ne $r.Nome }); 
            if ($dadosRestantes.Count -gt 0) { $dadosRestantes | Sort-Object Nome | Export-Csv $fileIPv6 -NoTypeInformation -Encoding UTF8 } else { "Nome,DNS_Primario,DNS_Secundario,Selecionado" | Out-File $fileIPv6 -Encoding UTF8 }; 
            Load-CsvToGrid 
        } 
    } 
})

# ==============================================================================
# 5.1 JANELA DE GERENCIAMENTO DA LISTA NEGRA (VERSÃO OTIMIZADA COM ABAS)
# ==============================================================================
function Show-BlacklistManager {
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Gerenciador da Lista Negra"
    $dlg.Size = New-Object System.Drawing.Size(550, 500) # Largura levemente aumentada para conforto
    $dlg.StartPosition = "CenterParent"
    $dlg.FormBorderStyle = "FixedDialog"
    $dlg.MaximizeBox = $false; $dlg.MinimizeBox = $false
    $dlg.BackColor = "White"

    # --- BARRA SUPERIOR (ABAS) ---
    $pnlTopBarBL = New-Object System.Windows.Forms.Panel
    $pnlTopBarBL.Dock = "Top"; $pnlTopBarBL.Height = 40; $pnlTopBarBL.BackColor = [System.Drawing.Color]::WhiteSmoke

    $btnTabBL4 = New-Object System.Windows.Forms.Button
    $btnTabBL4.Text = "Filtro IPv4"; $btnTabBL4.Size = New-Object System.Drawing.Size(130, 40); $btnTabBL4.Location = New-Object System.Drawing.Point(0, 0)
    $btnTabBL4.FlatStyle = "Flat"; $btnTabBL4.FlatAppearance.BorderSize = 0
    $btnTabBL4.BackColor = [System.Drawing.Color]::SteelBlue; $btnTabBL4.ForeColor = "White"; $btnTabBL4.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

    $btnTabBL6 = New-Object System.Windows.Forms.Button
    $btnTabBL6.Text = "Filtro IPv6"; $btnTabBL6.Size = New-Object System.Drawing.Size(130, 40); $btnTabBL6.Location = New-Object System.Drawing.Point(130, 0)
    $btnTabBL6.FlatStyle = "Flat"; $btnTabBL6.FlatAppearance.BorderSize = 0
    $btnTabBL6.BackColor = [System.Drawing.Color]::LightGray; $btnTabBL6.ForeColor = "Black"

    $pnlTopBarBL.Controls.Add($btnTabBL4); $pnlTopBarBL.Controls.Add($btnTabBL6)

    # --- BOTÕES DE AÇÃO (PAINEL INFERIOR) ---
    $pnlAcoes = New-Object System.Windows.Forms.Panel; $pnlAcoes.Dock = "Bottom"; $pnlAcoes.Height = 80

    $btnToggle = New-FlatButton "Ativar / Desativar" 25 20 150 "SlateGray"
    $btnRemover = New-FlatButton "Remover da Lista" 185 20 150 "IndianRed"
    $btnFechar = New-FlatButton "Salvar e Fechar" 345 20 160 "SteelBlue"
    $pnlAcoes.Controls.AddRange(@($btnToggle, $btnRemover, $btnFechar))

    # --- CONTAINERS DE CONTEÚDO (COM PADDING PARA NÃO COLAR NAS BORDAS) ---
    $pnlContentBL4 = New-Object System.Windows.Forms.Panel; $pnlContentBL4.Dock = "Fill"; $pnlContentBL4.Visible = $true
    $pnlContentBL4.Padding = New-Object System.Windows.Forms.Padding(15, 10, 15, 10)

    $pnlContentBL6 = New-Object System.Windows.Forms.Panel; $pnlContentBL6.Dock = "Fill"; $pnlContentBL6.Visible = $false
    $pnlContentBL6.Padding = New-Object System.Windows.Forms.Padding(15, 10, 15, 10)

    # --- FUNÇÃO HELPER PARA CRIAR OS GRIDS ---
    function New-BlacklistGrid {
        $grid = New-Object System.Windows.Forms.DataGridView
        $grid.Dock = "Fill" # Faz a tabela ocupar todo o espaço disponível no painel
        $grid.AutoSizeColumnsMode = "Fill"; $grid.AllowUserToAddRows = $false; $grid.ReadOnly = $true
        $grid.SelectionMode = "FullRowSelect"; $grid.MultiSelect = $false
        $grid.RowHeadersVisible = $false; $grid.BackgroundColor = "White"; $grid.BorderStyle = "FixedSingle"
        
        $grid.Add_DataBindingComplete({
            param($sender, $e)
            if ($sender.Columns.Contains("Status")) {
                $sender.Columns["Status"].AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::AllCells
                $sender.Columns["Status"].DefaultCellStyle.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleCenter
                $sender.Columns["Status"].DisplayIndex = 2 # Garante Status no final
            }
            if ($sender.Columns.Contains("Nome")) { $sender.Columns["Nome"].DisplayIndex = 0 }
            if ($sender.Columns.Contains("IP")) { $sender.Columns["IP"].DisplayIndex = 1 }
        })
        return $grid
    }

    $gridBL4 = New-BlacklistGrid; $pnlContentBL4.Controls.Add($gridBL4)
    $gridBL6 = New-BlacklistGrid; $pnlContentBL6.Controls.Add($gridBL6)

    # --- CARREGAMENTO E FILTRAGEM ---
    $blDados = @()
    if (Test-Path $fileBlacklist) { $blDados = @(Import-Csv $fileBlacklist -Encoding UTF8) }
    
    $listaIPv4 = [System.Collections.ArrayList]@($blDados | Where-Object { $_.IP -like "*.*" })
    $listaIPv6 = [System.Collections.ArrayList]@($blDados | Where-Object { $_.IP -like "*:*" })

    $gridBL4.DataSource = $listaIPv4
    $gridBL6.DataSource = $listaIPv6

    # --- LÓGICA DAS ABAS ---
    $btnTabBL4.Add_Click({
        $pnlContentBL4.Visible = $true; $pnlContentBL6.Visible = $false
        $btnTabBL4.BackColor = [System.Drawing.Color]::SteelBlue; $btnTabBL4.ForeColor = "White"; $btnTabBL4.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
        $btnTabBL6.BackColor = [System.Drawing.Color]::LightGray; $btnTabBL6.ForeColor = "Black"; $btnTabBL6.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    })
    $btnTabBL6.Add_Click({
        $pnlContentBL6.Visible = $true; $pnlContentBL4.Visible = $false
        $btnTabBL6.BackColor = [System.Drawing.Color]::SteelBlue; $btnTabBL6.ForeColor = "White"; $btnTabBL6.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
        $btnTabBL4.BackColor = [System.Drawing.Color]::LightGray; $btnTabBL4.ForeColor = "Black"; $btnTabBL4.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    })

    # --- LÓGICA DOS BOTÕES ---
    $btnToggle.Add_Click({
        $gridAtivo = if ($pnlContentBL4.Visible) { $gridBL4 } else { $gridBL6 }
        if ($gridAtivo.SelectedRows.Count -gt 0) {
            $r = $gridAtivo.SelectedRows[0].DataBoundItem
            $r.Status = if ($r.Status -eq "Ativo") { "Inativo" } else { "Ativo" }
            $gridAtivo.Refresh()
        }
    })

    $btnRemover.Add_Click({
        $gridAtivo = if ($pnlContentBL4.Visible) { $gridBL4 } else { $gridBL6 }
        if ($gridAtivo.SelectedRows.Count -gt 0) {
            $item = $gridAtivo.SelectedRows[0].DataBoundItem
            $lista = $gridAtivo.DataSource
            $lista.Remove($item)
            $gridAtivo.DataSource = $null; $gridAtivo.DataSource = $lista
        }
    })

    $btnFechar.Add_Click({
        $total = [System.Collections.ArrayList]::new()
        if ($null -ne $gridBL4.DataSource) { foreach ($i in $gridBL4.DataSource) { $total.Add($i) | Out-Null } }
        if ($null -ne $gridBL6.DataSource) { foreach ($i in $gridBL6.DataSource) { $total.Add($i) | Out-Null } }
        
        if ($total.Count -gt 0) {
            $total | Export-Csv $fileBlacklist -NoTypeInformation -Encoding UTF8
        } else {
            "IP,Nome,Status" | Out-File $fileBlacklist -Encoding UTF8
        }
        Load-CsvToGrid 
        $dlg.DialogResult = "OK"
    })

    # Ordem de adição controls (Z-Order): O que é adicionado por último fica "por trás" se usar Dock.Fill
    $dlg.Controls.Add($pnlContentBL4)
    $dlg.Controls.Add($pnlContentBL6)
    $dlg.Controls.Add($pnlAcoes)
    $dlg.Controls.Add($pnlTopBarBL)
    
    $dlg.ShowDialog() | Out-Null
    $dlg.Dispose()
}

# Conecta os botões roxos ao Gerenciador
$btnBlacklistIPv4.Add_Click({ Show-BlacklistManager })
$btnBlacklistIPv6.Add_Click({ Show-BlacklistManager })


# ==============================================================================
# 6. MOTOR DE TESTES (MULTITHREAD) E PROTEÇÃO DE BLACKLIST
# ==============================================================================
$tmrUI = New-Object System.Windows.Forms.Timer
$tmrUI.Interval = 50

$ScriptBlockTest = {
    param($sync, $servidores)
    
    $sync.IsRunning = $true
    $sync.Resultados.Clear()

    $sw = New-Object System.Diagnostics.Stopwatch

    foreach ($srv in $servidores) {
        if ($sync.Cancel) { break }
        $sync.MacroCount += 1
        $sync.MacroName = $srv.NomeShow
        $sync.MicroCount = 0

        $tempos = [System.Collections.ArrayList]::new()
        $falhou = $false

        for ($i = 1; $i -le 100; $i++) {
            if ($sync.Cancel) { break }
            $sync.MicroCount = $i
            try {
                $sw.Restart()
                $null = Resolve-DnsName -Name "google.com" -Server $srv.IP -DnsOnly -QuickTimeout -ErrorAction Stop
                $sw.Stop()
                $tempos.Add($sw.Elapsed.TotalMilliseconds) | Out-Null
            } catch {
                $falhou = $true
                break
            }
            Start-Sleep -Milliseconds 25
        }
        
        if ($sync.Cancel) { break }
        
        $rawMin = 0; $rawMax = 0; $rawMedia = 0; $rawMedMax = 0
        $strMin = "--------"; $strMax = "--------"; $strMedia = "--------"; $strMedMin = "--------"; $strMedMax = "--------"

        if (-not $falhou -and $tempos.Count -gt 0) {
            $rawMin = ($tempos | Measure-Object -Minimum).Minimum
            $rawMax = ($tempos | Measure-Object -Maximum).Maximum
            $rawMedia = ($tempos | Measure-Object -Average).Average
            
            $abaixo = $tempos | Where-Object { $_ -lt $rawMedia }
            $acima = $tempos | Where-Object { $_ -gt $rawMedia }
            
            if ($abaixo.Count -gt 0) { $strMedMin = "$([math]::Round(($abaixo | Measure-Object -Average).Average, 2)) ms" }
            if ($acima.Count -gt 0) { 
                $rawMedMax = ($acima | Measure-Object -Average).Average
                $strMedMax = "$([math]::Round($rawMedMax, 2)) ms" 
            }
            $strMin = "$([math]::Round($rawMin, 2)) ms"
            $strMax = "$([math]::Round($rawMax, 2)) ms"
            $strMedia = "$([math]::Round($rawMedia, 2)) ms"
        }

        $res = [PSCustomObject]@{
            Grupo = $srv.Grupo; Tipo = $srv.Tipo; Nome = $srv.NomeShow; IP = $srv.IP
            Mínimo = $strMin; 'Média-Mínima' = $strMedMin; Média = $strMedia
            'Média-Máxima' = $strMedMax; Máximo = $strMax
            RawMedia = $rawMedia; RawMedMax = $rawMedMax; RawMax = $rawMax
            TemposBrutos = $tempos.ToArray()
        }
        $sync.Resultados.Add($res) | Out-Null
    }
    $sync.IsRunning = $false
}

function Iniciar-Teste ($isIPv4) {
    $gridAtivo = if ($isIPv4) { $gridIPv4 } else { $gridIPv6 }
    $listaServidores = @()
    
    # Lê os dados de quem está marcado ("True" no DataBoundItem)
    foreach ($row in $gridAtivo.Rows) {
        $item = $row.DataBoundItem
        if ($row.Index -ge 0 -and $item.Selecionado -eq "True") {
            $nome = $item.Nome
            $pri = $item.DNS_Primario
            $sec = $item.DNS_Secundario
            
            if (-not [string]::IsNullOrWhiteSpace($pri) -and $global:BlacklistAtiva -notcontains $pri) { 
                $listaServidores += [PSCustomObject]@{ Grupo=$nome; Tipo="Primario"; NomeShow="$nome Primário"; IP=$pri } 
            }
            if (-not [string]::IsNullOrWhiteSpace($sec) -and $global:BlacklistAtiva -notcontains $sec) { 
                $listaServidores += [PSCustomObject]@{ Grupo=$nome; Tipo="Secundario"; NomeShow="$nome Secundário"; IP=$sec } 
            }
        }
    }
    
    # Trava de segurança caso o usuário tente iniciar o teste com tudo desmarcado
    if ($listaServidores.Count -eq 0) { 
        [System.Windows.Forms.MessageBox]::Show("Nenhum IP válido foi encontrado para o teste! Verifique as caixas selecionadas e a sua Lista Negra.", "Atenção", 0, "Warning")
        return 
    }

    $pnlMain.Visible = $false
    $pnlExec.Visible = $true
    
    $global:syncHash.Cancel = $false
    $global:syncHash.MacroCount = 0
    $global:syncHash.MicroCount = 0
    $global:syncHash.MacroName = "Preparando..."
    $global:syncHash.TotalMacro = $listaServidores.Count
    $pbMacro.Maximum = $listaServidores.Count
    $pbMacro.Value = 0

    $global:psInst = [System.Management.Automation.PowerShell]::Create()
    $global:psInst.AddScript($ScriptBlockTest).AddArgument($global:syncHash).AddArgument($listaServidores) | Out-Null
    $global:asyncResult = $global:psInst.BeginInvoke()
    $tmrUI.Start()
}

$btnTestIPv4.Add_Click({ Iniciar-Teste $true })
$btnTestIPv6.Add_Click({ Iniciar-Teste $false })

$btnCancelExec.Add_Click({
    $btnCancelExec.Text = "Cancelando..."
    $btnCancelExec.Enabled = $false
    $global:syncHash.Cancel = $true
})

# ==============================================================================
# 7. EXIBIÇÃO DE RESULTADOS E EXPORTAÇÃO
# ==============================================================================
$tmrUI.Add_Tick({
    $lblMacroExec.Text = "Testando: $($global:syncHash.MacroName) ($($global:syncHash.MacroCount)/$($global:syncHash.TotalMacro))"
    if ($global:syncHash.MacroCount -le $pbMacro.Maximum) { $pbMacro.Value = $global:syncHash.MacroCount }
    
    $lblMicroExec.Text = "Resoluções concluídas: $($global:syncHash.MicroCount)/100"
    if ($global:syncHash.MicroCount -le $pbMicro.Maximum) { $pbMicro.Value = $global:syncHash.MicroCount }

    if ($global:asyncResult.IsCompleted -or (-not $global:syncHash.IsRunning)) {
        $tmrUI.Stop()
        try { $global:psInst.EndInvoke($global:asyncResult) } catch {}
        
        $global:psInst.Dispose()
        $global:psInst = $null
        [System.GC]::Collect()

        $btnCancelExec.Text = "Cancelar Teste"
        $btnCancelExec.Enabled = $true

        if ($global:syncHash.Cancel) {
            $pnlExec.Visible = $false
            $pnlMain.Visible = $true
        } else {
            Preparar-TelaResultados
            $pnlExec.Visible = $false
            $pnlRes.Visible = $true
        }
    }
})

# --- MOTOR DE FILTROS ---
function Aplicar-FiltrosGrid {
    param([string]$tipo)
    $filtrados = [System.Collections.ArrayList]::new()

    foreach ($item in $global:syncHash.Resultados) {
        $isErro = ($item.Média -eq "--------")
        $isAlta = ($false)

        if (-not $isErro) {
            if ($item.RawMedia -ge 80 -or $item.RawMedMax -ge 80 -or $item.RawMax -ge 130) {
                $isAlta = $true
            }
        }

        if ($tipo -eq "Todos") { $filtrados.Add($item) | Out-Null }
        elseif ($tipo -eq "SemErro" -and -not $isErro) { $filtrados.Add($item) | Out-Null }
        elseif ($tipo -eq "SemCinza" -and -not $isErro -and -not $isAlta) { $filtrados.Add($item) | Out-Null }
        elseif ($tipo -eq "SemAltaLatencia" -and -not $isAlta) { $filtrados.Add($item) | Out-Null }
    }

    $gridRes.DataSource = $filtrados

    # Armadura de Erros
    if ($null -ne $gridRes.Columns["Grupo"]) { $gridRes.Columns["Grupo"].Visible = $false }
    if ($null -ne $gridRes.Columns["Tipo"]) { $gridRes.Columns["Tipo"].Visible = $false }
    if ($null -ne $gridRes.Columns["RawMedia"]) { $gridRes.Columns["RawMedia"].Visible = $false }
    if ($null -ne $gridRes.Columns["RawMedMax"]) { $gridRes.Columns["RawMedMax"].Visible = $false }
    if ($null -ne $gridRes.Columns["RawMax"]) { $gridRes.Columns["RawMax"].Visible = $false }
    if ($null -ne $gridRes.Columns["TemposBrutos"]) { $gridRes.Columns["TemposBrutos"].Visible = $false }
}

function Preparar-TelaResultados {
    # 1. Verifica os resultados para montar a Interface Dinâmica do Filtro
    $temErro = $false
    $temAlta = $false
    $temBom = $false

    foreach ($item in $global:syncHash.Resultados) {
        $isErro = ($item.Média -eq "--------")
        $isAlta = $false
        
        if (-not $isErro) {
            if ($item.RawMedia -ge 80 -or $item.RawMedMax -ge 80 -or $item.RawMax -ge 130) { 
                $isAlta = $true 
            }
        }

        if ($isErro) { $temErro = $true }
        elseif ($isAlta) { $temAlta = $true }
        else { $temBom = $true }
    }

    # O botão só liga se houver uma mistura de resultados para filtrar
    if (($temErro -or $temAlta) -and $temBom) {
        $btnFiltroRes.Enabled = $true
        $btnFiltroRes.BackColor = [System.Drawing.Color]::MediumPurple
    } else {
        $btnFiltroRes.Enabled = $false
        $btnFiltroRes.BackColor = [System.Drawing.Color]::LightGray
    }

    # 2. Constrói o menu suspenso sob medida para o cenário atual
    $ctxFiltro.Items.Clear()
    
    $itemTodos = $ctxFiltro.Items.Add("Mostrar Todos (Padrão)")
    $itemTodos.Checked = $true
    $itemTodos.Add_Click({ 
        foreach($i in $ctxFiltro.Items){ $i.Checked = $false }; $this.Checked = $true
        Aplicar-FiltrosGrid "Todos" 
    })
    
    # CENÁRIO 1: Deu ruim em tudo (Erros e Alta Latência)
    if ($temErro -and $temAlta) {
        $itemErro = $ctxFiltro.Items.Add("Ocultar Erros (Vermelhos)")
        $itemErro.Add_Click({ foreach($i in $ctxFiltro.Items){ $i.Checked = $false }; $this.Checked = $true; Aplicar-FiltrosGrid "SemErro" })
        
        $itemCombo = $ctxFiltro.Items.Add("Ocultar Alta Latência e Erros")
        $itemCombo.Add_Click({ foreach($i in $ctxFiltro.Items){ $i.Checked = $false }; $this.Checked = $true; Aplicar-FiltrosGrid "SemCinza" })
    }
    # CENÁRIO 2: Apenas Erros, sem Alta Latência
    elseif ($temErro -and -not $temAlta) {
        $itemErro = $ctxFiltro.Items.Add("Ocultar Erros (Vermelhos)")
        $itemErro.Add_Click({ foreach($i in $ctxFiltro.Items){ $i.Checked = $false }; $this.Checked = $true; Aplicar-FiltrosGrid "SemErro" })
    }
    # CENÁRIO 3: Apenas Alta Latência, sem Erros
    elseif (-not $temErro -and $temAlta) {
        $itemAlta = $ctxFiltro.Items.Add("Ocultar Alta Latência (Cinzas)")
        $itemAlta.Add_Click({ foreach($i in $ctxFiltro.Items){ $i.Checked = $false }; $this.Checked = $true; Aplicar-FiltrosGrid "SemAltaLatencia" })
    }

    # 3. Inicia a tela mostrando Todos os resultados
    Aplicar-FiltrosGrid "Todos"

    # 4. Matemática Original do Ranking
    $gruposDict = @{}
    foreach ($item in $global:syncHash.Resultados) {
        if ($item.Média -eq "--------") { continue }
        if (-not $gruposDict.ContainsKey($item.Grupo)) { $gruposDict[$item.Grupo] = @{} }
        $gruposDict[$item.Grupo][$item.Tipo] = $item
    }

    $ranking = @()
    foreach ($key in $gruposDict.Keys) {
        if ($gruposDict[$key].ContainsKey("Primario") -and $gruposDict[$key].ContainsKey("Secundario")) {
            $pri = $gruposDict[$key]["Primario"]; $sec = $gruposDict[$key]["Secundario"]
            $mediaComb = ($pri.RawMedia + $sec.RawMedia) / 2
            $medMaxComb = ($pri.RawMedMax + $sec.RawMedMax) / 2
            $maxComb = ($pri.RawMax + $sec.RawMax) / 2
            $pontos = ($mediaComb * 0.6) + ($medMaxComb * 0.2) + ($maxComb * 0.2)
            $ranking += [PSCustomObject]@{ Nome = $key; PriIP = $pri.IP; SecIP = $sec.IP; Score = $pontos }
        }
    }
    
    $ranking = $ranking | Sort-Object Score
    
    $global:rankingText = ""
    
    if ($ranking.Count -gt 0) {
        $blocoVencedor = "DNS com a melhor performance em geral:`nNome do dns       : $($ranking[0].Nome)`nDNS Primário      : $($ranking[0].PriIP)`nDNS Secundário    : $($ranking[0].SecIP)`n`n"
        $global:rankingText += $blocoVencedor
    }
    if ($ranking.Count -gt 1) {
        $blocoAlternativo = "DNS alternativa com a melhor performance em geral:`nNome do dns       : $($ranking[1].Nome)`nDNS Primário      : $($ranking[1].PriIP)`nDNS Secundário    : $($ranking[1].SecIP)"
        $global:rankingText += $blocoAlternativo
    } else {
        $global:rankingText += "Nenhuma alternativa disponível com ambos servidores respondendo sem falhas."
    }

    $txtResInfos.Text = @"
LEGENDA:
[VERDE]    BAIXA LATÊNCIA    (Média ≤60 ms e Média-Máxima <80 ms e Máximo <130 ms)
[BRANCO]   LATÊNCIA MODERADA (Média entre 61-79 ms, Média-Máxima <80 ms e Máximo <130 ms)
[CINZA]    ALTA LATÊNCIA     (Média ≥80 ms ou Média-Máxima ≥80 ms ou Máximo ≥130 ms)
[VERMELHO] ERRO              (Servidor ignorado por apresentar falha no teste)

$global:rankingText
"@
}

# ==============================================================================
# LÓGICA DE CAPTURA PARA A LISTA NEGRA
# ==============================================================================
$btnVoltarRes.Add_Click({ 
    # 1. Identifica quem deu erro. O uso do @() FORÇA o PowerShell a manter 
    # o formato de lista, mesmo se apenas 1 servidor falhar!
    $ipsComErro = @($global:syncHash.Resultados | Where-Object { $_.Média -eq "--------" })
    
    if ($ipsComErro.Count -gt 0) {
        $msg1 = "Alguns servidores falharam durante o teste. Deseja mover os DNS que apresentaram erro para a Lista Negra?"
        $resp1 = [System.Windows.Forms.MessageBox]::Show($msg1, "Lista Negra", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
        
        if ($resp1 -eq "Yes") {
            $msg2 = "Deseja adicionar TODOS os servidores que falharam ou ESCOLHER específicos?`n`n[ SIM ] = Adicionar TODOS`n[ NÃO ] = Escolher da lista"
            $resp2 = [System.Windows.Forms.MessageBox]::Show($msg2, "Opções de Adição", [System.Windows.Forms.MessageBoxButtons]::YesNoCancel, [System.Windows.Forms.MessageBoxIcon]::Question)
            
            $ipsParaAdicionar = @()
            
            if ($resp2 -eq "Yes") {
                # O usuário quis todos
                $ipsParaAdicionar = $ipsComErro
            } elseif ($resp2 -eq "No") {
                # O usuário quer escolher, abre a janelinha (CheckedListBox)
                $dlgPick = New-Object System.Windows.Forms.Form
                $dlgPick.Text = "Escolher DNS"
                $dlgPick.Size = New-Object System.Drawing.Size(350, 400)
                $dlgPick.StartPosition = "CenterParent"
                $dlgPick.FormBorderStyle = "FixedDialog"
                $dlgPick.MaximizeBox = $false; $dlgPick.MinimizeBox = $false
                $dlgPick.BackColor = "White"

                $lblPick = New-Object System.Windows.Forms.Label
                $lblPick.Text = "Marque os IPs que deseja isolar:"
                $lblPick.Location = New-Object System.Drawing.Point(15, 15)
                $lblPick.AutoSize = $true
                $lblPick.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
                
                $clb = New-Object System.Windows.Forms.CheckedListBox
                $clb.Location = New-Object System.Drawing.Point(15, 40)
                $clb.Size = New-Object System.Drawing.Size(300, 260)
                $clb.Font = New-Object System.Drawing.Font("Segoe UI", 10)
                $clb.CheckOnClick = $true
                
                # Coloca os IPs com erro na lista já marcados
                foreach ($err in $ipsComErro) {
                    $clb.Items.Add("$($err.Nome) - $($err.IP)", $true) | Out-Null
                }
                
                $btnConf = New-Object System.Windows.Forms.Button
                $btnConf.Text = "Confirmar"
                $btnConf.Location = New-Object System.Drawing.Point(115, 315)
                $btnConf.Size = New-Object System.Drawing.Size(100, 35)
                $btnConf.BackColor = [System.Drawing.Color]::SteelBlue; $btnConf.ForeColor = "White"
                $btnConf.FlatStyle = "Flat"; $btnConf.FlatAppearance.BorderSize = 0
                $btnConf.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
                $btnConf.DialogResult = "OK"
                
                $dlgPick.Controls.Add($lblPick); $dlgPick.Controls.Add($clb); $dlgPick.Controls.Add($btnConf)
                $dlgPick.AcceptButton = $btnConf
                
                if ($dlgPick.ShowDialog() -eq "OK") {
                    foreach ($itemChecked in $clb.CheckedItems) {
                        # Extrai o IP do texto e adiciona à lista (Blindado com @)
                        $ipExtraido = ($itemChecked -split " - ")[-1]
                        $ipsParaAdicionar += @($ipsComErro | Where-Object { $_.IP -eq $ipExtraido })
                    }
                }
                $dlgPick.Dispose()
            }
            
            # Pega os IPs escolhidos e salva no arquivo da Lista Negra
            if ($ipsParaAdicionar.Count -gt 0) {
                $blDados = @()
                if (Test-Path $fileBlacklist) { $blDados = @(Import-Csv $fileBlacklist -Encoding UTF8) }
                
                $novosAdicionados = $false
                foreach ($novo in $ipsParaAdicionar) {
                    # Só adiciona se o IP já não estiver na lista negra
                    $jaExiste = $false
                    if ($blDados.Count -gt 0) {
                        $jaExiste = [bool]($blDados | Where-Object { $_.IP -eq $novo.IP })
                    }
                    
                    if (-not $jaExiste) {
                        $novoObj = [PSCustomObject]@{ IP = $novo.IP; Nome = $novo.Nome; Status = "Ativo" }
                        $blDados += $novoObj
                        $novosAdicionados = $true
                    }
                }
                
                if ($novosAdicionados) {
                    $blDados | Export-Csv $fileBlacklist -NoTypeInformation -Encoding UTF8
                    Load-CsvToGrid
                }
            }
        }
    }
    
    # Transição final: Fecha os resultados e volta para a tela inicial
    $pnlRes.Visible = $false
    $pnlMain.Visible = $true 
})

$btnSaveRes.Add_Click({
    # Pega exatamente a lista que está visível na tela neste exato momento
    $listaParaSalvar = $gridRes.DataSource 
    
    # Verifica se há filtro ativo (se a tela tem menos itens que o total testado)
    $houveFiltro = ($listaParaSalvar.Count -lt $global:syncHash.Resultados.Count)

    if ($houveFiltro) {
        $msgFiltro = "Existem DNS lentos ou com erro ocultados neste momento.`n`nDeseja mostrar TODOS os servidores no arquivo de texto gerado?`n`n[ SIM ] = Salvar Todos (Ignorar Filtro)`n[ NÃO ] = Salvar Apenas Visíveis (Manter Filtro)"
        
        $resp = [System.Windows.Forms.MessageBox]::Show($msgFiltro, "Aviso de Filtro Ativo", [System.Windows.Forms.MessageBoxButtons]::YesNoCancel, [System.Windows.Forms.MessageBoxIcon]::Question)
        
        if ($resp -eq "Cancel") { return }
        if ($resp -eq "Yes") { 
            # Se disse sim, a lista que será salva volta a ser o resultado bruto total
            $listaParaSalvar = $global:syncHash.Resultados 
        }
    }

    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter = "Arquivo de Texto (*.txt)|*.txt"
    $sfd.FileName = "Resultado_DNS_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
    if ($sfd.ShowDialog() -eq "OK") {
        
        # 1. Encontra o maior tamanho usando a lista correta
        $maxNome = 30 
        $maxIP = 15   
        
        foreach ($item in $listaParaSalvar) {
            if ($item.Nome.Length -gt $maxNome) {
                $maxNome = $item.Nome.Length
            }
            if ($item.IP.Length -gt $maxIP) {
                $maxIP = $item.IP.Length
            }
        }
        
        $conteudo = @()
        
        # 2. Constrói o cabeçalho dinâmico
        $conteudo += ("Nome".PadRight($maxNome + 1) + "IP".PadRight($maxIP + 1) + "Mínimo".PadRight(13) + "Média-Mínima".PadRight(15) + "Média".PadRight(13) + "Média-Máxima".PadRight(15) + "Máximo")
        
        # 3. Constrói a linha tracejada dupla
        $linhaTracejada = ("-" * $maxNome) + " " + ("-" * $maxIP) + " ------------ -------------- ------------ -------------- ------------"
        $conteudo += $linhaTracejada
        
        # 4. Aplica o preenchimento dinâmico nas linhas que vão ser salvas
        foreach ($item in $listaParaSalvar) {
            $conteudo += $item.Nome.PadRight($maxNome) + " " + $item.IP.PadRight($maxIP) + " " +
                         $item.Mínimo.PadRight(12) + " " + $item.('Média-Mínima').PadRight(14) + " " +
                         $item.Média.PadRight(12) + " " + $item.('Média-Máxima').PadRight(14) + " " + $item.Máximo
        }
        
        $conteudo += "`n`n$($global:rankingText)"
        
        $conteudo | Out-File -FilePath $sfd.FileName -Encoding UTF8
        [System.Windows.Forms.MessageBox]::Show("Salvo com sucesso!", "Salvar", 0, "Information")
    }
    $sfd.Dispose()
})

# ==============================================================================
# 8. INICIALIZAÇÃO
# ==============================================================================
Load-CsvToGrid
$form.Add_FormClosed({ 
    $tmrUI.Dispose()
    if ($null -ne $ctxMenu) { $ctxMenu.Dispose() }
    if ($null -ne $ctxFiltro) { $ctxFiltro.Dispose() }
    
    # --- Libera a chave do Mutex no sistema ---
    if ($null -ne $global:appMutex) { 
        $global:appMutex.ReleaseMutex()
        $global:appMutex.Dispose() 
    }
    
    # --- Limpeza definitiva do Ícone (Blinda contra Memory Leak GDI) ---
    if ($null -ne $form.Icon) { $form.Icon.Dispose() }
    if ($null -ne $global:hIconGlobe -and $global:hIconGlobe -ne [IntPtr]::Zero) { 
        [Win32]::DestroyIcon($global:hIconGlobe) | Out-Null
    }
    
    $form.Dispose()
    [System.GC]::Collect() 
})
$form.ShowDialog() | Out-Null