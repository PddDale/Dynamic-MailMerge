# Disparo interativo de e-mails via Outlook

*[Read this in English](README.md)*

Script Python que automatiza o envio de e-mails personalizados em massa
usando o Outlook clássico (via automação COM/MAPI), a partir de uma
planilha de contatos. É genérico: detecta as contas e caixas
compartilhadas configuradas no perfil do Outlook do usuário e pergunta
interativamente tudo que precisa saber — não há nada fixo para uma
empresa específica.

Existem duas versões equivalentes do script, cada uma com prompts,
comentários e nomes de colunas do log no seu próprio idioma:

- `MailMerge_Dinamico.py` — versão em português (documentada aqui).
- `MailMerge_Dynamic.py` — versão em inglês (veja o README em inglês).

As duas rodam exatamente o mesmo fluxo; use a que preferir.

## O que o script faz

1. **Verifica o ambiente**: confere se as dependências Python estão
   instaladas (instala automaticamente as que faltarem) e se o Outlook
   clássico está aberto (tenta abrir automaticamente se não estiver, e
   avisa caso detecte o "Novo Outlook", que não suporta essa automação).
2. **Identifica o usuário**: lista os e-mails já configurados como conta
   no Outlook para escolha rápida (por número) ou aceita digitar um
   e-mail manualmente. Usado apenas para registro no log.
3. **Seleciona a caixa de envio**: lista as contas completas e as
   caixas compartilhadas/pastas adicionais detectadas no perfil. Se a
   escolha for uma caixa compartilhada, pergunta se o envio deve ser
   feito "em nome de" ela (resolvendo o endereço no catálogo de
   endereços do Exchange/GAL). Se preferir não usar essa permissão, o
   envio segue pela conta padrão do Outlook.
4. **Seleciona a planilha de contatos**: pede o caminho de um arquivo
   Excel (.xlsx/.xls), deixa escolher a aba e as colunas de nome e
   e-mail, e avisa sobre linhas vazias ou células com múltiplos valores
   separados por "/" (que são puladas no envio, pois quebrariam a
   personalização individual).
5. **Assinatura (Word) e logo**: pede o caminho de um documento Word
   (.docx) com o padrão de assinatura. A conversão para HTML é feita
   diretamente a partir das runs do `python-docx`, preservando negrito,
   itálico, sublinhado, cor exata do texto, tamanho de fonte e imagens
   já embutidas no documento (com a largura/altura originais do Word).
   Se o documento já tiver uma imagem embutida, a etapa de logo externa
   é pulada automaticamente para evitar duplicidade; caso contrário,
   pergunta se há uma logo separada (PNG/JPG) para anexar — nesse caso
   a imagem é redimensionada proporcionalmente (máx. 180×60px) para não
   ficar desproporcional na assinatura.
6. **Template do corpo do e-mail**: pede o caminho de um template
   (.txt, .html ou .docx) contendo apenas o corpo específico da
   mensagem — sem saudação e sem assinatura. A saudação "Prezado(a)
   `<Nome>`," é gerada automaticamente a partir da planilha, e a
   assinatura é sempre a extraída na etapa anterior. Templates `.docx`
   também têm a formatação (negrito, cor, tamanho de fonte, imagens)
   preservada.
7. **Assunto e confirmação de leitura**: pede o assunto do e-mail e
   pergunta se deve solicitar recibo de leitura ao destinatário.
8. **Log de envio**: pergunta se deve ser gerado um log de envio
   (arquivo Excel com o status de cada e-mail). Se sim, pergunta o
   caminho onde salvar (Enter usa o padrão: `log_envio.xlsx` na mesma
   pasta da planilha de contatos).
9. **Resumo e confirmação**: mostra um resumo de tudo que foi
   configurado — incluindo se o log de envio será gerado e em qual
   caminho — e pede confirmação final antes de iniciar o envio.
10. **Envio**: dispara um e-mail por contato válido da planilha, com
    pausa entre envios. Se a geração de log foi confirmada, salva o
    arquivo no caminho escolhido com o status de cada linha (enviado,
    erro, ou pulado por dados inválidos).

## Verificação de confirmação de leitura

O script também tem um segundo modo de operação, para depois de um
disparo já feito: aponta para o `log_envio.xlsx` gerado e escolhe a
caixa usada no envio original a partir da mesma lista de contas e
caixas compartilhadas/pastas adicionais detectadas no perfil do Outlook
(mesmo sistema usado na seleção da caixa de envio, na Etapa 3 — também
aceita digitar o e-mail/nome manualmente). O script então varre a
caixa de entrada correspondente procurando os recibos de leitura
(confirmados ou recusados), atualizando o log com o status e a data/hora
de cada leitura no formato `DD/MM/AAAA HH:MM`.

Quando o envio original foi feito "em nome de" uma caixa compartilhada,
o Exchange costuma entregar os recibos de leitura na Caixa de Entrada
da conta pessoal que de fato enviou a mensagem, e não na Caixa de
Entrada própria da caixa compartilhada. Por isso, além da caixa
escolhida, o script sempre verifica também a Caixa de Entrada padrão
do usuário atual, evitando que os recibos passem despercebidos (e sem
duplicar resultados, caso as duas apontem para a mesma pasta).

## Requisitos

- Windows com Outlook clássico instalado e configurado (o "Novo
  Outlook" não suporta a automação COM/MAPI usada aqui).
- Python 3, com as bibliotecas `pywin32`, `pandas`, `openpyxl`,
  `python-docx` e `Pillow` (instaladas automaticamente pelo script na
  Etapa 1, se ainda não estiverem presentes).

## Como usar

```
python MailMerge_Dinamico.py
```

O script é totalmente interativo: basta seguir as perguntas exibidas no
terminal, na ordem descrita acima.
