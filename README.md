![Outlook Logo](https://outlookiniciarsesion01.weebly.com/uploads/9/8/5/4/98549006/outlook_orig.jpg)

# Processo de Envio de E-mails com Anexos em Massa / Outlook

Este script em Python automatiza o envio de e-mails em massa utilizando o servidor SMTP do Outlook. Ele obtém informações de destinatários e anexos de um arquivo Excel e envia e-mails personalizados com anexos PDF. Abaixo está uma descrição detalhada do funcionamento do script:

## Bibliotecas Utilizadas

- `smtplib`: Envio de e-mails através do protocolo SMTP.
- `email.mime`: Criação e manipulação de mensagens de e-mail.
- `os`: Interações com o sistema operacional, como verificação de arquivos.
- `openpyxl`: Leitura de arquivos Excel.

## Configurações do Servidor SMTP

- **Servidor SMTP**: `smtp-mail.outlook.com`
- **Porta**: 587 (TLS)

## Informações de Login

- **E-mail do Usuário**: Deve ser inserido no campo `email_usuario`.
- **Senha**: Deve ser inserida no campo `senha`.

## Caminho do Arquivo Excel

- **Arquivo Excel**: O caminho do arquivo Excel contendo os dados dos destinatários deve ser especificado em `caminho_arquivo_excel`.

## Função para Obter Destinatários e Anexos

A função `obter_destinatarios_e_anexos` lê o arquivo Excel e obtém os endereços de e-mail e caminhos dos arquivos PDF a serem anexados. O formato do caminho do PDF é corrigido e verificado para garantir que o arquivo existe.

## Função para Enviar o E-mail

A função `enviar_email` monta e envia o e-mail para cada destinatário. O e-mail inclui:

- **Corpo do E-mail**: Texto personalizado com o nome do destinatário.
- **Anexo PDF**: Arquivo PDF cujo caminho é especificado no arquivo Excel.

### Passos para Envio do E-mail

1. **Criação da Mensagem**: Utiliza `MIMEMultipart` para criar a mensagem de e-mail.
2. **Corpo do E-mail**: Adiciona um texto personalizado ao corpo do e-mail.
3. **Anexar Arquivo PDF**: Lê e anexa o arquivo PDF.
4. **Conexão com o Servidor SMTP**: Conecta-se ao servidor SMTP do Outlook, autentica e envia o e-mail.
5. **Tratamento de Erros**: Inclui um bloco `try-except` para lidar com possíveis erros durante o envio.

## Execução do Script

A execução principal do script chama a função `obter_destinatarios_e_anexos` para obter os destinatários e seus respectivos anexos, e então itera sobre essa lista para enviar os e-mails.

### Exemplo de Execução

```python
if __name__ == "__main__":
    destinatarios_e_anexos = obter_destinatarios_e_anexos()
    for destinatario, nome, caminho_anexo_pdf in destinatarios_e_anexos:
        enviar_email(destinatario, nome, caminho_anexo_pdf)
