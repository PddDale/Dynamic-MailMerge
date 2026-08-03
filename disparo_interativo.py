# -*- coding: utf-8 -*-
"""
Assistente interativo de disparo de e-mails personalizados via Outlook
clássico, usando a caixa compartilhada rfp.bts@ereadvisory.com como
remetente (via SentOnBehalfOfName + resolução no GAL do Exchange).

Este script mantém integralmente a lógica de resolução da caixa
compartilhada validada em "Script de Envio.py" (ver HANDOFF_disparo_emails.md,
gotchas 2 e 3): SentOnBehalfOfName + CreateRecipient() com verificação de
AddressEntry.Type == "EX" e fallback pelo nome de exibição no GAL.

Uso: python disparo_interativo.py
"""

import glob
import importlib
import os
import re
import subprocess
import sys
import time

# ==================== CONFIGURAÇÃO FIXA (NÃO ALTERAR) ====================
# A caixa de envio real (remetente, via SentOnBehalfOfName) é sempre esta,
# independentemente do e-mail que o usuário informar na Etapa 2. A Etapa 2
# é puramente informativa/log de quem disparou a campanha.
CONTA_ENVIO = "rfp.bts@ereadvisory.com"
NOME_EXIBICAO_FALLBACK = "rfp bts"  # nome de exibição no GAL, usado como fallback
# ===========================================================================

def limpar_caminho(texto):
    """Aceita o caminho colado 'cru', entre aspas ("..."/'...') ou entre
    parênteses ((...)) — formatos comuns ao copiar caminhos do Explorer ou
    de outras fontes — e devolve só o caminho limpo."""
    texto = texto.strip()
    pares = [('"', '"'), ("'", "'"), ("(", ")")]
    for abre, fecha in pares:
        if len(texto) >= 2 and texto.startswith(abre) and texto.endswith(fecha):
            texto = texto[1:-1].strip()
    return texto


BLOCO_ASSINATURA_PADRAO = """<p>Atenciosamente,</p>
<p><b>Priscila Milfano</b><br>
<p><b>Expansão & Gestão de Portfólio</b><br>
Cell BR: +55 11 94003-6409</p>
<p>Condomínio Edifício Aliança<br>
Rua Verbo Divino, 1547 | 13º andar<br>
São Paulo/SP &ndash; CEP 04719-002<br>
CRECI 32378-J</p>
<p><img src="cid:logo_assinatura" width="185" height="48"></p>"""


# ==================== ETAPA 1: VERIFICAÇÃO DE AMBIENTE ====================

def verificar_dependencias():
    pacotes = {
        "win32com.client": "pywin32",
        "pandas": "pandas",
        "openpyxl": "openpyxl",
        "docx": "python-docx",
    }
    faltando = []
    for modulo, pacote_pip in pacotes.items():
        try:
            importlib.import_module(modulo)
        except ImportError:
            faltando.append(pacote_pip)

    if faltando:
        print(f"[Ambiente] Instalando pacotes ausentes: {', '.join(faltando)}...")
        subprocess.run([sys.executable, "-m", "pip", "install", *faltando], check=True)
        print(f"[Ambiente] Instalado(s): {', '.join(faltando)}")
    else:
        print("[Ambiente] pywin32, pandas e openpyxl já estão instalados.")


def outlook_classico_em_execucao():
    try:
        saida = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq OUTLOOK.EXE"],
            capture_output=True, text=True, check=False,
        )
        return "OUTLOOK.EXE" in saida.stdout.upper()
    except Exception:
        return False


def localizar_outlook_exe():
    candidatos = []
    for base_env in ("PROGRAMFILES", "PROGRAMFILES(X86)"):
        base = os.environ.get(base_env)
        if not base:
            continue
        candidatos.extend(glob.glob(os.path.join(base, "Microsoft Office", "root", "Office*", "OUTLOOK.EXE")))
        candidatos.extend(glob.glob(os.path.join(base, "Microsoft Office", "Office*", "OUTLOOK.EXE")))
    for caminho in candidatos:
        if os.path.exists(caminho):
            return caminho
    return None


def detectar_novo_outlook():
    """Heurística: o "Novo Outlook" roda como app empacotado (olk.exe /
    WindowsApps), não como OUTLOOK.EXE clássico. Se só existir processo de
    app empacotado e nenhum OUTLOOK.EXE, alertamos o usuário."""
    try:
        saida = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq olk.exe"],
            capture_output=True, text=True, check=False,
        )
        return "OLK.EXE" in saida.stdout.upper()
    except Exception:
        return False


def verificar_outlook_classico():
    if detectar_novo_outlook() and not outlook_classico_em_execucao():
        print("\n[ATENÇÃO] Foi detectado o 'Novo Outlook' em execução, que NÃO suporta")
        print("automação COM/MAPI. Por favor, feche o Novo Outlook, abra o Outlook")
        print("CLÁSSICO e pressione Enter para continuar.")
        input()
        return

    if outlook_classico_em_execucao():
        print("[Ambiente] Outlook clássico já está em execução.")
        return

    print("[Ambiente] Outlook clássico não detectado. Tentando abrir automaticamente...")
    exe = localizar_outlook_exe()
    if exe:
        try:
            subprocess.Popen([exe])
            print(f"[Ambiente] Outlook iniciado a partir de: {exe}")
            print("[Ambiente] Aguardando alguns segundos para a automação COM ficar disponível...")
            time.sleep(8)
            return
        except Exception as e:
            print(f"[Ambiente] Falha ao iniciar o Outlook automaticamente: {e}")

    print("\n[AÇÃO NECESSÁRIA] Não foi possível localizar/abrir o Outlook clássico")
    print("automaticamente. Por favor, abra o Outlook clássico manualmente e")
    print("pressione Enter para continuar.")
    input()


def executar_verificacao_ambiente():
    print("=" * 70)
    print("ETAPA 1/5 - Verificação de ambiente")
    print("=" * 70)
    verificar_dependencias()
    verificar_outlook_classico()
    print()


# ==================== ETAPA 2: E-MAIL DO USUÁRIO (INFORMATIVO) =============

def perguntar_email_usuario():
    print("=" * 70)
    print("ETAPA 2/5 - Identificação do usuário (apenas para log)")
    print("=" * 70)
    regex_email = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")
    while True:
        email = input("Qual é o seu e-mail principal (ex: john.doe@ereadvisory.com)? ").strip()
        if regex_email.match(email):
            print(f"[Info] E-mail do usuário registrado para log: {email}")
            print(f"[Info] O remetente real do disparo continua sendo: {CONTA_ENVIO} (fixo, não afetado por essa resposta)")
            print()
            return email
        print("E-mail inválido. Tente novamente.")


# ==================== ETAPA 3: SELEÇÃO DA PLANILHA =========================

def perguntar_caminho_planilha():
    import pandas as pd

    while True:
        caminho = limpar_caminho(input("Caminho do arquivo Excel (.xlsx/.xls) com os contatos: "))
        if not os.path.exists(caminho):
            print("Arquivo não encontrado. Tente novamente.")
            continue
        if not caminho.lower().endswith((".xlsx", ".xls")):
            print("O arquivo precisa ser .xlsx ou .xls. Tente novamente.")
            continue
        try:
            xls = pd.ExcelFile(caminho)
            return caminho, xls
        except Exception as e:
            print(f"Não foi possível ler o arquivo ({e}). Tente novamente.")


def escolher_aba(xls):
    abas = xls.sheet_names
    while True:
        print("\nAbas disponíveis:")
        for i, nome in enumerate(abas, start=1):
            print(f"  {i}. {nome}")
        escolha = input("Escolha a aba (número ou nome): ").strip()
        if escolha.isdigit() and 1 <= int(escolha) <= len(abas):
            return abas[int(escolha) - 1]
        if escolha in abas:
            return escolha
        print("Escolha inválida. Tente novamente.")


def listar_colunas_preenchidas(df):
    colunas_validas = [c for c in df.columns if str(c).strip() and not str(c).startswith("Unnamed:")]
    print("\nColunas disponíveis:")
    for i, col in enumerate(colunas_validas, start=1):
        print(f"  {i}. {col}")
    return colunas_validas


def escolher_coluna(df, rotulo):
    while True:
        colunas_validas = listar_colunas_preenchidas(df)
        escolha = input(f"Qual coluna contém os {rotulo}? (número ou nome): ").strip()
        if escolha.isdigit() and 1 <= int(escolha) <= len(colunas_validas):
            return colunas_validas[int(escolha) - 1]
        if escolha in colunas_validas:
            return escolha
        print("Escolha inválida. Tente novamente.")


def validar_coluna_dados(df, coluna, rotulo):
    serie = df[coluna]
    vazias = serie.isna() | (serie.astype(str).str.strip() == "")
    n_vazias = int(vazias.sum())
    if n_vazias == len(serie):
        print(f"[AVISO] A coluna '{coluna}' escolhida para {rotulo} está totalmente vazia.")
    elif n_vazias > 0:
        print(f"[AVISO] A coluna '{coluna}' ({rotulo}) tem {n_vazias} linha(s) vazia(s).")

    multiplos = serie.dropna().astype(str).apply(lambda v: "/" in v)
    n_multiplos = int(multiplos.sum())
    if n_multiplos > 0:
        print(f"[AVISO] A coluna '{coluna}' ({rotulo}) tem {n_multiplos} célula(s) com múltiplos valores")
        print("        separados por '/'. Essas linhas precisam ser explodidas em linhas")
        print("        separadas antes do envio para manter a personalização individual")
        print("        (ver HANDOFF_disparo_emails.md, item 6). Elas serão puladas no envio")
        print("        automático se não forem corrigidas na planilha.")


def selecionar_planilha():
    print("=" * 70)
    print("ETAPA 3/5 - Seleção da planilha de contatos")
    print("=" * 70)
    import pandas as pd

    caminho, xls = perguntar_caminho_planilha()
    aba = escolher_aba(xls)
    df = pd.read_excel(caminho, sheet_name=aba)

    print(f"\nAba '{aba}' selecionada. Agora escolha as colunas de nome e e-mail.")
    coluna_nome = escolher_coluna(df, "NOMES")
    coluna_email = escolher_coluna(df, "E-MAILS")

    validar_coluna_dados(df, coluna_nome, "nomes")
    validar_coluna_dados(df, coluna_email, "e-mails")
    print()

    return {
        "caminho": caminho,
        "aba": aba,
        "df": df,
        "coluna_nome": coluna_nome,
        "coluna_email": coluna_email,
    }


# ==================== ETAPA 4: IMAGEM DE ASSINATURA =========================

def perguntar_caminho_logo():
    print("=" * 70)
    print("ETAPA 4/5 - Imagem de assinatura")
    print("=" * 70)
    while True:
        caminho = limpar_caminho(input("Caminho do arquivo de imagem da logo/assinatura (PNG/JPG): "))
        if os.path.exists(caminho) and caminho.lower().endswith((".png", ".jpg", ".jpeg")):
            print()
            return caminho
        print("Arquivo não encontrado ou formato inválido (use .png/.jpg/.jpeg). Tente novamente.")


# ==================== ETAPA 5: TEMPLATE DO CORPO ============================
# O template contém APENAS o corpo específico da mensagem (sem saudação e
# sem assinatura). A saudação "Prezado(a) <Nome>," é gerada automaticamente
# pelo script a partir da coluna de nomes da planilha, e a assinatura é
# sempre o bloco fixo BLOCO_ASSINATURA_PADRAO (nunca muda, nunca vem do
# template).

def extrair_texto_docx(caminho):
    from docx import Document

    doc = Document(caminho)
    paragrafos = []
    for para in doc.paragraphs:
        texto = para.text.strip()
        if texto:
            paragrafos.append(f"<p>{texto}</p>")
    return "\n".join(paragrafos)


def extrair_corpo_texto_ou_html(caminho):
    with open(caminho, "r", encoding="utf-8") as f:
        texto = f.read()
    if caminho.lower().endswith((".html", ".htm")):
        return texto
    linhas = texto.replace("\r\n", "\n").split("\n")
    return "\n".join(f"<p>{linha.strip()}</p>" for linha in linhas if linha.strip())


def envolver_em_html(texto):
    if "<html" in texto.lower():
        return texto
    return (
        '<html>\n<body style="font-family:Calibri, Arial, sans-serif; font-size:11pt; color:#000000;">\n'
        f"{texto}\n</body>\n</html>"
    )


def selecionar_template():
    print("=" * 70)
    print("ETAPA 5/5 - Template do corpo do e-mail")
    print("=" * 70)
    print("O template deve conter apenas o corpo específico da mensagem (sem")
    print("saudação inicial e sem assinatura) — a saudação é gerada")
    print("automaticamente a partir do nome na planilha, e a assinatura é")
    print("sempre o modelo fixo da empresa.")
    while True:
        caminho = limpar_caminho(input("Caminho do template do corpo do e-mail (.txt, .html ou .docx): "))
        if not os.path.exists(caminho) or not caminho.lower().endswith((".txt", ".html", ".htm", ".docx")):
            print("Arquivo não encontrado ou formato inválido (use .txt/.html/.docx). Tente novamente.")
            continue
        if os.path.getsize(caminho) == 0:
            print("O arquivo está vazio (0 bytes). Abra-o, adicione o conteúdo e salve novamente. Tente outro caminho.")
            continue
        try:
            if caminho.lower().endswith(".docx"):
                corpo = extrair_texto_docx(caminho)
            else:
                corpo = extrair_corpo_texto_ou_html(caminho)
        except Exception as e:
            print(f"Não foi possível ler o arquivo ({e}). Tente novamente.")
            continue
        if not corpo.strip():
            print("O template está vazio. Tente novamente.")
            continue
        break

    corpo_completo = f"<p>Prezado(a) {{nome}},</p>\n{corpo}\n{BLOCO_ASSINATURA_PADRAO}"
    html_final = envolver_em_html(corpo_completo)
    print()
    return caminho, html_final


# ==================== RESOLUÇÃO DA CAIXA COMPARTILHADA ======================
# Lógica mantida exatamente como validada em Script de Envio.py / HANDOFF
# (gotchas 2 e 3): resolver via CreateRecipient(), checar AddressEntry.Type
# == "EX"; se não for, tentar de novo pelo nome de exibição no GAL.

def resolver_caixa_compartilhada(namespace):
    def resolver(nome_ou_email):
        rec = namespace.CreateRecipient(nome_ou_email)
        rec.Resolve()
        if not rec.Resolved:
            return None
        tipo = rec.AddressEntry.Type
        print(f"  Tentativa '{nome_ou_email}' -> Resolvido | Tipo: {tipo} | Address: {rec.Address}")
        return rec, tipo

    print("Resolvendo caixa compartilhada contra o catálogo de endereços...")
    resultado = resolver(CONTA_ENVIO)

    if resultado is None or resultado[1] != "EX":
        print(f"Endereço via e-mail não é nativo do Exchange (ou não resolveu). "
              f"Tentando pelo nome de exibição '{NOME_EXIBICAO_FALLBACK}'...")
        resultado = resolver(NOME_EXIBICAO_FALLBACK)

    if resultado is None:
        raise Exception(
            f"Não foi possível resolver {CONTA_ENVIO} nem '{NOME_EXIBICAO_FALLBACK}' no catálogo de endereços. "
            "Confirme o nome exato como aparece na busca do GAL."
        )

    rec, tipo = resultado
    if tipo != "EX":
        print(f"\nATENÇÃO: o endereço resolvido é do tipo '{tipo}', não 'EX' (Exchange nativo).")
        print("Isso provavelmente vai continuar dando erro de permissão.\n")

    return rec.Address


# ==================== RESUMO E CONFIRMAÇÃO ===================================

def contar_linhas_validas(df, coluna_nome, coluna_email):
    validos = df[coluna_nome].notna() & (df[coluna_nome].astype(str).str.strip() != "") \
        & df[coluna_email].notna() & (df[coluna_email].astype(str).str.strip() != "")
    return int(validos.sum())


def mostrar_resumo_e_confirmar(email_usuario, planilha, logo_path, template_path, assunto):
    n_linhas = contar_linhas_validas(planilha["df"], planilha["coluna_nome"], planilha["coluna_email"])
    print("=" * 70)
    print("RESUMO DA CONFIGURAÇÃO")
    print("=" * 70)
    print(f"Disparado por (log):      {email_usuario}")
    print(f"Caixa de envio (fixo):    {CONTA_ENVIO}")
    print(f"Assunto:                  {assunto}")
    print(f"Planilha:                 {planilha['caminho']}")
    print(f"Aba:                      {planilha['aba']}")
    print(f"Coluna de nomes:          {planilha['coluna_nome']}")
    print(f"Coluna de e-mails:        {planilha['coluna_email']}")
    print(f"Logo/assinatura:          {logo_path}")
    print(f"Template:                 {template_path}")
    print(f"Linhas válidas a enviar:  {n_linhas}")
    print("=" * 70)
    resposta = input("Confirma o início do envio? (s/n): ").strip().lower()
    return resposta == "s"


# ==================== ENVIO ===================================================

def enviar_emails(outlook, endereco_resolvido, planilha, logo_path, corpo_html_template, assunto):
    df = planilha["df"]
    coluna_nome = planilha["coluna_nome"]
    coluna_email = planilha["coluna_email"]

    import pandas as pd

    log = []
    for _, row in df.iterrows():
        nome_raw = row[coluna_nome]
        email_raw = row[coluna_email]

        if pd.isna(nome_raw) or pd.isna(email_raw):
            continue
        nome = str(nome_raw).strip()
        email = str(email_raw).strip()
        if not nome or not email:
            continue
        if "/" in nome or "/" in email:
            log.append({"Nome": nome, "Email": email, "Status": "Pulado: múltiplos valores na célula"})
            print(f"[PULADO] {nome} / {email}: célula com múltiplos valores, precisa ser explodida antes.")
            continue

        try:
            mail = outlook.CreateItem(0)  # 0 = olMailItem
            mail.To = email
            mail.Subject = assunto
            mail.HTMLBody = corpo_html_template.replace("{nome}", nome)

            attachment = mail.Attachments.Add(logo_path)
            attachment.PropertyAccessor.SetProperty(
                "http://schemas.microsoft.com/mapi/proptag/0x3712001E",
                "logo_assinatura"
            )

            mail.SentOnBehalfOfName = endereco_resolvido
            mail.Send()

            log.append({"Nome": nome, "Email": email, "Status": "Enviado"})
            print(f"[OK] Enviado para {nome} ({email})")

        except Exception as e:
            log.append({"Nome": nome, "Email": email, "Status": f"Erro: {e}"})
            print(f"[ERRO] Falha ao enviar para {email}: {e}")

        time.sleep(2)

    log_df = pd.DataFrame(log)
    log_path = os.path.join(os.path.dirname(planilha["caminho"]), "log_envio.xlsx")
    log_df.to_excel(log_path, index=False)
    print(f"\nConcluído. Log salvo em: {log_path}")


# ==================== ORQUESTRAÇÃO ============================================

def main():
    executar_verificacao_ambiente()
    email_usuario = perguntar_email_usuario()
    planilha = selecionar_planilha()
    logo_path = perguntar_caminho_logo()
    template_path, corpo_html_template = selecionar_template()

    assunto = input("Assunto do e-mail: ").strip()

    if not mostrar_resumo_e_confirmar(email_usuario, planilha, logo_path, template_path, assunto):
        print("Operação cancelada pelo usuário.")
        return

    import win32com.client as win32

    outlook = win32.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    endereco_resolvido = resolver_caixa_compartilhada(namespace)
    print(f"\nEndereço final usado no envio: {endereco_resolvido}")
    print(f"Enviando em nome de: {CONTA_ENVIO}\n")

    enviar_emails(outlook, endereco_resolvido, planilha, logo_path, corpo_html_template, assunto)


if __name__ == "__main__":
    main()