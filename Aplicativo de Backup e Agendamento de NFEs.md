# Aplicativo de Backup e Agendamento de NFEs

Este é um aplicativo Python com interface gráfica (Tkinter) para automatizar a cópia, compactação e upload de arquivos XML (NFEs) para o Google Drive, com notificações por e-mail e agendamento de tarefas.

## Requisitos

Para executar este aplicativo, você precisará:

1.  **Python 3.x**: Certifique-se de ter o Python 3 instalado em seu sistema. Você pode baixá-lo em [python.org](https://www.python.org/downloads/).
2.  **Tkinter**: A biblioteca Tkinter geralmente vem com a instalação padrão do Python. Se você estiver em um sistema Linux e tiver problemas, pode precisar instalá-la separadamente (ex: `sudo apt-get install python3-tk`).
3.  **rclone**: Este aplicativo utiliza o `rclone` para o upload de arquivos para o Google Drive. Baixe e configure o `rclone` em seu sistema. Certifique-se de que o caminho para o executável `rclone.exe` e o arquivo de configuração `rclone.conf` estejam corretos nas configurações do aplicativo.
    *   Baixe o rclone em [rclone.org/downloads/](https://rclone.org/downloads/)
    *   Configure um remote para o Google Drive (ex: `MeuGoogleDrive`) seguindo a documentação do rclone.

## Como Usar

1.  **Baixe os arquivos do aplicativo**: Copie o arquivo `app.py` para uma pasta em seu computador.

2.  **Execute o aplicativo**: Abra um terminal ou prompt de comando na pasta onde você salvou o `app.py` e execute o seguinte comando:
    ```bash
    python app.py
    ```

3.  **Configurações (Aba "Configurações")**:
    *   **Pasta Origem**: Onde os arquivos XML (NFEs) estão localizados.
    *   **Pasta Destino Base**: Onde os arquivos serão copiados localmente antes da compactação e upload.
    *   **Caminho Rclone**: O caminho completo para o executável `rclone.exe`.
    *   **Nome Remoto Rclone**: O nome do remote do rclone configurado para o Google Drive (ex: `MeuGoogleDrive`).
    *   **Pasta Base Drive**: A pasta raiz no seu Google Drive onde os backups serão armazenados (ex: `CLIENTES`).
    *   **Nome Cliente Específico**: O nome do cliente para organizar as pastas no Google Drive (ex: `PIE HOUSE`).
    *   **Configurações de E-mail**: Preencha os detalhes do seu servidor SMTP, usuário, senha (senha de aplicativo para Gmail) e endereços de e-mail para notificações.
    *   **Opções de Funcionalidades**: Marque ou desmarque as caixas para ativar/desativar notificações por e-mail, upload para o Google Drive via rclone e verificação de pré-requisitos.

4.  **Execução (Aba "Execução")**:
    *   Clique no botão "Executar Backup Agora" para iniciar o processo de cópia, compactação e upload manualmente.
    *   A área de texto exibirá os logs de execução em tempo real.

5.  **Agendamento (Aba "Agendamento")**:
    *   Clique no botão "Agendar Tarefa" para criar uma tarefa agendada no Windows. Esta tarefa executará o backup automaticamente todo dia 1 de cada mês às 10:00.
    *   **Importante**: O agendamento de tarefas no Windows requer permissões de administrador. Execute o aplicativo como administrador se o agendamento falhar.
    *   O aplicativo tentará criar um script PowerShell temporário (`run_backup.ps1`) na sua `Pasta Destino Base` para ser usado pela tarefa agendada. Este script chamará o aplicativo Python para executar o backup.

## Observações Importantes

*   **Segurança da Senha**: A senha do SMTP é armazenada em texto simples no aplicativo. Para maior segurança em um ambiente de produção, considere usar métodos mais seguros de armazenamento de credenciais (ex: variáveis de ambiente, gerenciadores de segredos).
*   **Caminhos no Windows**: Certifique-se de usar barras invertidas duplas (`\\`) ou barras normais (`/`) em caminhos de arquivo no Python para evitar problemas de escape (ex: `C:\\Mobility_POS\\Xml_IO`).
*   **Execução como Administrador**: Para agendar tarefas no Windows, o aplicativo (ou o terminal de onde ele é executado) precisa ter permissões de administrador.
*   **Logs**: Os logs de execução são exibidos na interface e também salvos em um arquivo `log_copia_nfe.log` na sua `Pasta Destino Base`.

## Solução de Problemas

*   **`ModuleNotFoundError: No module named 'tkinter'`**: Instale o Tkinter para Python. Em Debian/Ubuntu: `sudo apt-get install python3-tk`.
*   **`_tkinter.TclError: no display name and no $DISPLAY environment variable`**: Este erro ocorre em ambientes sem interface gráfica (como servidores). O aplicativo precisa de um ambiente gráfico para ser executado.
*   **Problemas com rclone**: Verifique se o `rclone.exe` está no caminho correto e se o remote (`MeuGoogleDrive`) está configurado corretamente (`rclone config`). Verifique também o arquivo `rclone.conf`.
*   **Problemas de E-mail**: Verifique as configurações do servidor SMTP, porta, usuário e senha. Para Gmail, você precisará de uma "senha de aplicativo" (App Password) se tiver a verificação em duas etapas ativada.

---

Espero que este aplicativo seja útil!

