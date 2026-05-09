# Monitor de Latência DNS

**Descubra os servidores DNS mais rápidos para a sua conexão. Utilitário completo com interface amigável, ranking de velocidade e sistema de exclusão.**

O **Monitor de Latência DNS** é uma ferramenta robusta desenvolvida em PowerShell com interface gráfica (WinForms) projetada para entusiastas de rede e usuários que buscam otimizar sua conexão de internet. Através de testes multithread precisos, a aplicação avalia a performance de diversos servidores DNS simultaneamente, ajudando você a escolher a melhor rota para sua navegação ou jogos.

---

## 🚀 Funcionalidades

- **Interface Gráfica Nativa:** UI fluida e responsiva com aceleração gráfica (Double Buffering) para evitar cintilação.
- **Suporte Dual-Stack:** Gerenciamento e testes completos para servidores **IPv4** e **IPv6**.
- **Testes Multithread:** Realiza resoluções DNS em segundo plano, permitindo que a interface permaneça responsiva durante o processo.
- **Ranking de Performance:** Algoritmo inteligente que gera um ranking dos melhores servidores baseando-se na média, picos e estabilidade.
- **Sistema de Lista Negra (Blacklist):** Filtre IPs instáveis ou sem resposta para que sejam ignorados em testes futuros.
- **Exportação de Dados:** Salve os resultados detalhados em arquivos `.txt` formatados.
- **Instalação Simplificada:** Inclui uma lista de servidores DNS já montada e pronta para uso.

---

## 📸 Capturas de Tela

Aqui você pode visualizar a interface do programa em operação:

### 1. Tela Principal
Interface de gerenciamento onde você pode selecionar, adicionar e editar seus servidores DNS preferidos.
![Tela Principal](main.png)

### 2. Execução de Testes
Monitoramento em tempo real com barras de progresso macro e micro, exibindo o status individual de cada servidor.
![Tela de Teste](test_screen.png)

### 3. Gerenciamento de Lista Negra
Controle total sobre IPs isolados por mau desempenho ou falta de resposta.
![Lista Negra](blacklist.png)

---

## 📦 Como Usar

1. **Script PowerShell:** Baixe o arquivo `.ps1` e execute-o com o botão direito -> *Run with PowerShell*.
2. **Servidores DNS:** Para facilitar, este repositório inclui um arquivo zipado com uma lista de servidores DNS já montada. Basta extrair na pasta do script.
3. **Versão Executável:** Se preferir não lidar com scripts, disponibilizei uma versão já convertida para `.exe` na seção de [Releases/Arquivos].

---

## 🛠️ Compilação Manual (PowerShell para EXE)

Se você for um desenvolvedor e desejar converter o script `.ps1` em um executável `.exe` por conta própria, recomenda-se o uso da ferramenta **ps2exe**.

Para garantir que a interface gráfica e o suporte a caracteres especiais funcionem corretamente, os seguintes parâmetros são **obrigatórios**:

```powershell
ps2exe.ps1 -inputFile .\Monitor_de_Latencia.ps1 -outputFile .\Monitor_de_Latencia.exe -supportOS -sta -noConsole -unicodeEncoding
