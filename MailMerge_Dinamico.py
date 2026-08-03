# -*- coding: utf-8 -*-
"""
Assistente interativo de disparo de e-mails personalizados via Outlook
clássico.

O script detecta as contas e caixas compartilhadas disponíveis no perfil do
Outlook do usuário, pergunta qual delas deve ser usada como remetente e, se
for uma caixa compartilhada, oferece a opção de enviar "em nome de" ela
(via SentOnBehalfOfName + resolução no catálogo de endereços do Exchange).
Também pergunta um documento Word com o padrão de assinatura (a formatação
do Word é preservada ao converter para HTML) e, opcionalmente, uma imagem de
logo para incluir na assinatura.

Uso: python disparo_interativo.py
"""

import glob
import importlib
import os
import re
import subprocess
import sys
import time


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


# ==================== ETAPA 1: VERIFICAÇÃO DE AMBIENTE ====================

def verificar_dependencias():
    pacotes = {
        "win32com.client": "pywin32",
        "pandas": "pandas",
        "openpyxl": "openpyxl",
        "docx": "python-docx",
        "PIL": "Pillow",
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
        print("[Ambiente] Todas as dependências já estão instaladas.")


def _listar_processos_em_execucao():
    """Roda o 'tasklist' uma única vez (sem filtro) para que as duas
    verificações de processo (Outlook clássico e Novo Outlook) não precisem
    cada uma disparar seu próprio processo tasklist."""
    try:
        saida = subprocess.run(
            ["tasklist"],
            capture_output=True, text=True, check=False,
        )
        return saida.stdout.upper()
    except Exception:
        return ""


def outlook_classico_em_execucao(lista_processos=None):
    if lista_processos is None:
        lista_processos = _listar_processos_em_execucao()
    return "OUTLOOK.EXE" in lista_processos


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


def detectar_novo_outlook(lista_processos=None):
    """Heurística: o "Novo Outlook" roda como app empacotado (olk.exe /
    WindowsApps), não como OUTLOOK.EXE clássico. Se só existir processo de
    app empacotado e nenhum OUTLOOK.EXE, alertamos o usuário."""
    if lista_processos is None:
        lista_processos = _listar_processos_em_execucao()
    return "OLK.EXE" in lista_processos


def verificar_outlook_classico():
    lista_processos = _listar_processos_em_execucao()
    if detectar_novo_outlook(lista_processos) and not outlook_classico_em_execucao(lista_processos):
        print("\n[ATENÇÃO] Foi detectado o 'Novo Outlook' em execução, que NÃO suporta")
        print("automação COM/MAPI. Por favor, feche o Novo Outlook, abra o Outlook")
        print("CLÁSSICO e pressione Enter para continuar.")
        input()
        return

    if outlook_classico_em_execucao(lista_processos):
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
    print("ETAPA 1/6 - Verificação de ambiente")
    print("=" * 70)
    verificar_dependencias()
    verificar_outlook_classico()
    print()


# ==================== ETAPA 2: E-MAIL DO USUÁRIO (INFORMATIVO) =============

def perguntar_email_usuario(caixas_disponiveis):
    print("=" * 70)
    print("ETAPA 2/6 - Identificação do usuário (apenas para log)")
    print("=" * 70)
    regex_email = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")

    contas = [c["rotulo"] for c in caixas_disponiveis if c["tipo"] == "conta"]
    if contas:
        print("Emails configurados:")
        for i, email in enumerate(contas, start=1):
            print(f"  {i}. {email}")
        while True:
            escolha = input("Qual deve ser usado (número da lista ou digite o e-mail): ").strip()
            if escolha.isdigit() and 1 <= int(escolha) <= len(contas):
                email = contas[int(escolha) - 1]
                print(f"[Info] E-mail do usuário registrado para log: {email}")
                print()
                return email
            if regex_email.match(escolha):
                print(f"[Info] E-mail do usuário registrado para log: {escolha}")
                print()
                return escolha
            print("Escolha inválida. Digite o número de uma opção da lista ou um e-mail válido.")

    while True:
        email = input("Qual é o seu e-mail principal (ex: john.doe@company.com)? ").strip()
        if regex_email.match(email):
            print(f"[Info] E-mail do usuário registrado para log: {email}")
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

    multiplos = serie.dropna().astype(str).str.contains("/", regex=False)
    n_multiplos = int(multiplos.sum())
    if n_multiplos > 0:
        print(f"[AVISO] A coluna '{coluna}' ({rotulo}) tem {n_multiplos} célula(s) com múltiplos valores")
        print("        separados por '/'. Essas linhas precisam ser explodidas em linhas")
        print("        separadas antes do envio para manter a personalização individual.")
        print("        Elas serão puladas no envio automático se não forem corrigidas na planilha.")


def selecionar_planilha():
    print("=" * 70)
    print("ETAPA 4/6 - Seleção da planilha de contatos")
    print("=" * 70)

    caminho, xls = perguntar_caminho_planilha()
    aba = escolher_aba(xls)
    df = xls.parse(sheet_name=aba)

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


# ==================== ETAPA 4: SELEÇÃO DA CAIXA DE ENVIO ====================
# O script detecta as contas completas e as caixas compartilhadas/pastas
# adicionais disponíveis no perfil do Outlook e pergunta ao usuário qual
# delas deve ser usada como remetente. Se a escolha for uma caixa
# compartilhada (não uma conta completa do próprio usuário), pergunta se o
# envio deve ser feito "em nome de" ela (SentOnBehalfOfName), resolvendo o
# endereço no catálogo de endereços do Exchange (GAL). Se o usuário preferir
# não usar essa permissão, o envio segue pela conta padrão do Outlook.

def listar_caixas_disponiveis(namespace):
    caixas = []
    try:
        for acc in namespace.Accounts:
            try:
                email = acc.SmtpAddress
            except Exception:
                continue
            if email:
                caixas.append({"rotulo": email, "tipo": "conta", "conta_obj": acc})
    except Exception:
        pass

    rotulos_conhecidos = {c["rotulo"].lower() for c in caixas}
    try:
        for folder in namespace.Folders:
            try:
                nome = (folder.Name or "").strip()
            except Exception:
                continue
            if nome and nome.lower() not in rotulos_conhecidos:
                caixas.append({"rotulo": nome, "tipo": "pasta_adicional", "conta_obj": None})
                rotulos_conhecidos.add(nome.lower())
    except Exception:
        pass

    return caixas


def caixa_configurada_no_perfil(caixas_disponiveis, alvo):
    """Verifica se 'alvo' (e-mail ou nome de exibição) aparece registrado
    como conta completa ou como pasta adicional (caixa compartilhada/Full
    Access) no perfil atual do Outlook, reaproveitando a lista de caixas já
    obtida em vez de percorrer o COM do Outlook de novo."""
    alvo_lower = alvo.strip().lower()

    for caixa in caixas_disponiveis:
        rotulo_lower = caixa["rotulo"].strip().lower()
        if caixa["tipo"] == "conta":
            if rotulo_lower == alvo_lower:
                return True
        elif alvo_lower in rotulo_lower:
            return True

    return False


def resolver_endereco_envio(namespace, nome_ou_email_inicial):
    def resolver(valor):
        rec = namespace.CreateRecipient(valor)
        rec.Resolve()
        if not rec.Resolved:
            return None
        tipo = rec.AddressEntry.Type
        print(f"  Tentativa '{valor}' -> Resolvido | Tipo: {tipo} | Address: {rec.Address}")
        return rec, tipo

    print("Resolvendo endereço de envio contra o catálogo de endereços...")
    resultado = resolver(nome_ou_email_inicial)

    if resultado is None or resultado[1] != "EX":
        nome_fallback = input(
            f"Não foi possível resolver '{nome_ou_email_inicial}' como endereço nativo do Exchange.\n"
            "Digite o nome de exibição exato como aparece na busca do catálogo de endereços "
            "(ou apenas Enter para tentar de novo com o mesmo valor): "
        ).strip()
        if nome_fallback:
            resultado = resolver(nome_fallback)

    if resultado is None:
        raise Exception(
            f"Não foi possível resolver '{nome_ou_email_inicial}' no catálogo de endereços. "
            "Confirme o nome exato como aparece na busca do catálogo (GAL)."
        )

    rec, tipo = resultado
    if tipo != "EX":
        print(f"\nATENÇÃO: o endereço resolvido é do tipo '{tipo}', não 'EX' (Exchange nativo).")
        print("Isso pode causar erro de permissão ao usar 'Enviar em nome de'.\n")

    return rec.Address


def perguntar_caixa_envio(namespace, caixas):
    print("=" * 70)
    print("ETAPA 3/6 - Seleção da caixa de envio")
    print("=" * 70)

    print("Caixas detectadas no perfil do Outlook:")
    for i, caixa in enumerate(caixas, start=1):
        rotulo_tipo = "conta completa" if caixa["tipo"] == "conta" else "pasta adicional / possível caixa compartilhada"
        print(f"  {i}. {caixa['rotulo']} ({rotulo_tipo})")

    while True:
        escolha = input("Qual caixa deve ser usada como remetente (número da lista ou digite o e-mail/nome): ").strip()
        if escolha.isdigit() and 1 <= int(escolha) <= len(caixas):
            caixa_escolhida = caixas[int(escolha) - 1]
            break
        if escolha:
            caixa_escolhida = {"rotulo": escolha, "tipo": "manual", "conta_obj": None}
            break
        print("Valor vazio. Digite o número de uma opção da lista ou um e-mail/nome válido.")

    if caixa_escolhida["tipo"] == "conta":
        print(f"\n'{caixa_escolhida['rotulo']}' é uma conta completa configurada no perfil.")
        print("O envio usará essa conta diretamente, sem necessidade de 'Enviar em nome de'.\n")
        return {"modo": "conta_completa", "endereco_envio": caixa_escolhida["rotulo"], "conta_obj": caixa_escolhida["conta_obj"]}

    print(f"\nVerificando se '{caixa_escolhida['rotulo']}' está configurada no perfil...")
    if caixa_configurada_no_perfil(caixas, caixa_escolhida["rotulo"]):
        print("Caixa compartilhada/pasta adicional detectada no perfil.")
    else:
        print("Não foi detectada como pasta adicional/conta no perfil (ainda pode funcionar via")
        print("permissão 'Enviar em nome de', desde que ela tenha sido concedida no Exchange).")

    resposta = input(
        f"Deseja enviar os e-mails em nome de '{caixa_escolhida['rotulo']}' (via 'Enviar em nome de')? (s/n): "
    ).strip().lower()
    if resposta != "s":
        print("Ok, o envio usará a conta padrão do Outlook (sem 'Enviar em nome de').\n")
        return {"modo": "conta_completa", "endereco_envio": None, "conta_obj": None}

    endereco_resolvido = resolver_endereco_envio(namespace, caixa_escolhida["rotulo"])
    print(f"\nEndereço final resolvido para o envio: {endereco_resolvido}\n")
    return {"modo": "caixa_compartilhada", "endereco_envio": endereco_resolvido, "conta_obj": None}


# ==================== ETAPA 5: ASSINATURA (DOCX) E LOGO ======================
# O padrão de assinatura é fornecido pelo usuário como um documento Word. A
# conversão para HTML é feita diretamente a partir das runs do python-docx
# (não usa uma biblioteca de simplificação como o mammoth), para preservar
# fielmente negrito, itálico, sublinhado, cor exata do texto e imagens
# embutidas no próprio documento (ex.: logo já colada na assinatura).

_ALINHAMENTO_CSS = {
    1: "center",  # WD_ALIGN_PARAGRAPH.CENTER
    2: "right",   # WD_ALIGN_PARAGRAPH.RIGHT
    3: "justify",  # WD_ALIGN_PARAGRAPH.JUSTIFY
}


_EMU_POR_PIXEL = 9525  # EMU por pixel a 96 DPI (padrão do Word/Office)


def _extrair_imagens_run(run):
    import base64
    from docx.oxml.ns import qn

    imagens_html = []
    for drawing in run._element.findall(".//" + qn("w:drawing")):
        largura_px = altura_px = None
        extent = drawing.find(".//" + qn("wp:extent"))
        if extent is not None:
            cx = extent.get("cx")
            cy = extent.get("cy")
            if cx and cy:
                try:
                    largura_px = round(int(cx) / _EMU_POR_PIXEL)
                    altura_px = round(int(cy) / _EMU_POR_PIXEL)
                except Exception:
                    largura_px = altura_px = None

        for blip in drawing.findall(".//" + qn("a:blip")):
            r_id = blip.get(qn("r:embed"))
            if not r_id:
                continue
            try:
                image_part = run.part.related_parts[r_id]
                b64 = base64.b64encode(image_part.blob).decode("ascii")
                dimensoes = f' width="{largura_px}" height="{altura_px}"' if largura_px and altura_px else ""
                imagens_html.append(f'<img src="data:{image_part.content_type};base64,{b64}"{dimensoes}>')
            except Exception:
                continue
    return imagens_html


def _run_para_html(run):
    from html import escape

    partes = []
    texto = run.text or ""
    if texto:
        texto_html = escape(texto).replace("\n", "<br>")
        estilos = []
        cor = None
        try:
            if run.font.color is not None and run.font.color.type is not None and run.font.color.rgb is not None:
                cor = str(run.font.color.rgb)
        except Exception:
            pass
        if cor:
            estilos.append(f"color:#{cor}")
        try:
            if run.font.size is not None:
                estilos.append(f"font-size:{run.font.size.pt}pt")
        except Exception:
            pass
        if run.bold:
            texto_html = f"<strong>{texto_html}</strong>"
        if run.italic:
            texto_html = f"<em>{texto_html}</em>"
        if run.underline:
            texto_html = f"<u>{texto_html}</u>"
        if estilos:
            texto_html = f'<span style="{";".join(estilos)}">{texto_html}</span>'
        partes.append(texto_html)
    partes.extend(_extrair_imagens_run(run))
    return "".join(partes)


def _extrair_paragrafos_docx_html(caminho):
    """Converte cada parágrafo de um .docx para HTML preservando negrito,
    itálico, sublinhado, cor e tamanho de fonte de cada trecho de texto, além
    de imagens embutidas (com a largura/altura originais do documento)."""
    from docx import Document

    doc = Document(caminho)
    paragrafos_html = []
    for paragrafo in doc.paragraphs:
        conteudo = "".join(_run_para_html(run) for run in paragrafo.runs)
        if not conteudo.strip():
            continue
        alinhamento = _ALINHAMENTO_CSS.get(paragrafo.alignment)
        estilo = f' style="text-align:{alinhamento}"' if alinhamento else ""
        paragrafos_html.append(f"<p{estilo}>{conteudo}</p>")
    return paragrafos_html


def extrair_assinatura_docx(caminho):
    return "\n".join(_extrair_paragrafos_docx_html(caminho))


def perguntar_caminho_assinatura():
    print("=" * 70)
    print("ETAPA 5/6 - Padrão de assinatura (Word) e logo")
    print("=" * 70)
    while True:
        caminho = limpar_caminho(input("Caminho do documento Word (.docx) com o padrão de assinatura: "))
        if not os.path.exists(caminho) or not caminho.lower().endswith(".docx"):
            print("Arquivo não encontrado ou formato inválido (use .docx). Tente novamente.")
            continue
        if os.path.getsize(caminho) == 0:
            print("O arquivo está vazio (0 bytes). Abra-o, adicione o conteúdo e salve novamente. Tente outro caminho.")
            continue
        try:
            assinatura_html = extrair_assinatura_docx(caminho)
        except Exception as e:
            print(f"Não foi possível ler o documento ({e}). Tente novamente.")
            continue
        if not assinatura_html.strip():
            print("O documento de assinatura está vazio. Tente novamente.")
            continue
        return caminho, assinatura_html


def assinatura_ja_tem_imagem(assinatura_html):
    return "<img" in assinatura_html


def perguntar_logo(assinatura_html):
    if assinatura_ja_tem_imagem(assinatura_html):
        print("[Info] O documento de assinatura já contém uma imagem embutida (extraída com o")
        print("tamanho original do Word). A etapa de logo externa foi pulada para evitar duplicidade.")
        print()
        return None

    resposta = input("Existe uma imagem de logo da empresa para incluir na assinatura? (s/n): ").strip().lower()
    if resposta != "s":
        print()
        return None
    while True:
        caminho = limpar_caminho(input("Caminho do arquivo de imagem da logo (PNG/JPG): "))
        if os.path.exists(caminho) and caminho.lower().endswith((".png", ".jpg", ".jpeg")):
            print()
            return caminho
        print("Arquivo não encontrado ou formato inválido (use .png/.jpg/.jpeg). Tente novamente.")


LARGURA_MAXIMA_LOGO_PX = 180
ALTURA_MAXIMA_LOGO_PX = 60


def calcular_dimensoes_logo(caminho, largura_max=LARGURA_MAXIMA_LOGO_PX, altura_max=ALTURA_MAXIMA_LOGO_PX):
    """Lê o tamanho real da imagem e devolve dimensões reduzidas
    proporcionalmente para caber dentro do limite (nunca aumenta a imagem,
    só evita que uma logo em alta resolução apareça gigante na assinatura)."""
    from PIL import Image

    try:
        with Image.open(caminho) as img:
            largura, altura = img.size
    except Exception:
        return None
    if largura <= 0 or altura <= 0:
        return None

    fator = min(1.0, largura_max / largura, altura_max / altura)
    return max(1, round(largura * fator)), max(1, round(altura * fator))


def montar_assinatura_final(assinatura_html, logo_path):
    if logo_path:
        dimensoes = calcular_dimensoes_logo(logo_path)
        atributos = f' width="{dimensoes[0]}" height="{dimensoes[1]}"' if dimensoes else ""
        return f'{assinatura_html}\n<p><img src="cid:logo_assinatura"{atributos}></p>'
    return assinatura_html


# ==================== ETAPA 6: TEMPLATE DO CORPO ============================
# O template contém APENAS o corpo específico da mensagem (sem saudação e
# sem assinatura). A saudação "Prezado(a) <Nome>," é gerada automaticamente
# pelo script a partir da coluna de nomes da planilha, e a assinatura é
# sempre a extraída do documento Word na etapa anterior.

def extrair_texto_docx(caminho):
    return "\n".join(_extrair_paragrafos_docx_html(caminho))


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


def selecionar_template(assinatura_html):
    print("=" * 70)
    print("ETAPA 6/6 - Template do corpo do e-mail")
    print("=" * 70)
    print("O template deve conter apenas o corpo específico da mensagem (sem")
    print("saudação inicial e sem assinatura) — a saudação é gerada")
    print("automaticamente a partir do nome na planilha, e a assinatura é")
    print("sempre a extraída do documento Word informado na etapa anterior.")
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

    corpo_completo = f"<p>Prezado(a) {{nome}},</p>\n{corpo}\n{assinatura_html}"
    html_final = envolver_em_html(corpo_completo)
    print()
    return caminho, html_final


# ==================== RESUMO E CONFIRMAÇÃO ===================================

def contar_linhas_validas(df, coluna_nome, coluna_email):
    validos = df[coluna_nome].notna() & (df[coluna_nome].astype(str).str.strip() != "") \
        & df[coluna_email].notna() & (df[coluna_email].astype(str).str.strip() != "")
    return int(validos.sum())


def perguntar_confirmacao_leitura():
    print("=" * 70)
    print("Confirmação de leitura")
    print("=" * 70)
    print("O Outlook pode solicitar ao destinatário um recibo de leitura (o")
    print("destinatário ainda pode recusar o envio do recibo). Depois do disparo,")
    print("use a opção 'Verificar confirmações de leitura' deste script, apontando")
    print("para o log gerado, para atualizar o status de cada contato.")
    resposta = input("Solicitar confirmação de leitura para este disparo? (s/n): ").strip().lower()
    return resposta == "s"


def perguntar_geracao_log(planilha):
    print("=" * 70)
    print("Log de envio")
    print("=" * 70)
    resposta = input("Deseja gerar um log de envio (arquivo Excel) com o status de cada e-mail? (s/n): ").strip().lower()
    if resposta != "s":
        print()
        return False, None

    caminho_padrao = os.path.join(os.path.dirname(planilha["caminho"]), "log_envio.xlsx")
    caminho = limpar_caminho(input(
        f"Caminho onde o log deve ser salvo (Enter para usar o padrão: {caminho_padrao}): "
    ))
    if not caminho:
        print()
        return True, caminho_padrao

    if os.path.isdir(caminho):
        caminho = os.path.join(caminho, "log_envio.xlsx")
    elif not caminho.lower().endswith((".xlsx", ".xls")):
        caminho = caminho + ".xlsx"
    print()
    return True, caminho


def mostrar_resumo_e_confirmar(email_usuario, planilha, caixa_envio, logo_path, template_path, assunto,
                                solicitar_confirmacao_leitura, gerar_log, caminho_log):
    n_linhas = contar_linhas_validas(planilha["df"], planilha["coluna_nome"], planilha["coluna_email"])
    rotulo_caixa = caixa_envio["endereco_envio"] or "conta padrão do Outlook"
    print("=" * 70)
    print("RESUMO DA CONFIGURAÇÃO")
    print("=" * 70)
    print(f"Disparado por (log):      {email_usuario}")
    print(f"Caixa de envio:           {rotulo_caixa}")
    print(f"Modo de envio:            {'Enviar em nome de (caixa compartilhada)' if caixa_envio['modo'] == 'caixa_compartilhada' else 'Conta direta'}")
    print(f"Assunto:                  {assunto}")
    print(f"Planilha:                 {planilha['caminho']}")
    print(f"Aba:                      {planilha['aba']}")
    print(f"Coluna de nomes:          {planilha['coluna_nome']}")
    print(f"Coluna de e-mails:        {planilha['coluna_email']}")
    print(f"Logo na assinatura:       {logo_path or 'Nenhuma'}")
    print(f"Template:                 {template_path}")
    print(f"Linhas válidas a enviar:  {n_linhas}")
    print(f"Confirmação de leitura:   {'Sim' if solicitar_confirmacao_leitura else 'Não'}")
    print(f"Log de envio:             {'Sim (' + caminho_log + ')' if gerar_log else 'Não'}")
    print("=" * 70)
    resposta = input("Confirma o início do envio? (s/n): ").strip().lower()
    return resposta == "s"


# ==================== ENVIO ===================================================

def enviar_emails(outlook, caixa_envio, planilha, logo_path, corpo_html_template, assunto,
                   solicitar_confirmacao_leitura=False, gerar_log=True, caminho_log=None):
    df = planilha["df"]
    coluna_nome = planilha["coluna_nome"]
    coluna_email = planilha["coluna_email"]

    import pandas as pd
    from datetime import datetime

    status_leitura_inicial = "Pendente" if solicitar_confirmacao_leitura else "Não solicitada"

    log = []
    for nome_raw, email_raw in zip(df[coluna_nome], df[coluna_email]):
        if pd.isna(nome_raw) or pd.isna(email_raw):
            continue
        nome = str(nome_raw).strip()
        email = str(email_raw).strip()
        if not nome or not email:
            continue
        if "/" in nome or "/" in email:
            log.append({
                "Nome": nome, "Email": email, "Assunto": assunto, "Status": "Pulado: múltiplos valores na célula",
                "DataHoraEnvio": "", "ConfirmacaoLeitura": "N/A", "DataHoraLeitura": "",
            })
            print(f"[PULADO] {nome} / {email}: célula com múltiplos valores, precisa ser explodida antes.")
            continue

        try:
            mail = outlook.CreateItem(0)  # 0 = olMailItem
            mail.To = email
            mail.Subject = assunto
            mail.HTMLBody = corpo_html_template.replace("{nome}", nome)

            if logo_path:
                attachment = mail.Attachments.Add(logo_path)
                attachment.PropertyAccessor.SetProperty(
                    "http://schemas.microsoft.com/mapi/proptag/0x3712001E",
                    "logo_assinatura"
                )

            if caixa_envio["conta_obj"] is not None:
                mail.SendUsingAccount = caixa_envio["conta_obj"]
            elif caixa_envio["modo"] == "caixa_compartilhada" and caixa_envio["endereco_envio"]:
                mail.SentOnBehalfOfName = caixa_envio["endereco_envio"]

            if solicitar_confirmacao_leitura:
                mail.ReadReceiptRequested = True
            mail.Send()

            log.append({
                "Nome": nome, "Email": email, "Assunto": assunto, "Status": "Enviado",
                "DataHoraEnvio": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "ConfirmacaoLeitura": status_leitura_inicial, "DataHoraLeitura": "",
            })
            print(f"[OK] Enviado para {nome} ({email})")

        except Exception as e:
            log.append({
                "Nome": nome, "Email": email, "Assunto": assunto, "Status": f"Erro: {e}",
                "DataHoraEnvio": "", "ConfirmacaoLeitura": "N/A", "DataHoraLeitura": "",
            })
            print(f"[ERRO] Falha ao enviar para {email}: {e}")

        time.sleep(2)

    if not gerar_log:
        print("\nConcluído. Geração de log de envio desativada nesta execução.")
        return None

    log_df = pd.DataFrame(log)
    log_path = caminho_log or os.path.join(os.path.dirname(planilha["caminho"]), "log_envio.xlsx")
    log_df.to_excel(log_path, index=False)
    print(f"\nConcluído. Log salvo em: {log_path}")
    if solicitar_confirmacao_leitura:
        print("Recibos de leitura solicitados. Use a opção 'Verificar confirmações de leitura'")
        print(f"deste script apontando para o arquivo acima para atualizar o status.")
    return log_path


# ==================== VERIFICAÇÃO DE CONFIRMAÇÃO DE LEITURA =================
# O Outlook, quando ReadReceiptRequested=True é definido no envio, faz o
# destinatário receber um pedido de recibo de leitura; se ele aceitar, um
# item de relatório (MessageClass "REPORT.IPM.Note.IPNRN" para confirmado ou
# "REPORT.IPM.Note.IPNNRN" para recusado — em algumas configurações pode vir
# sem o prefixo "REPORT.") chega na Caixa de Entrada da conta
# que enviou. Esta seção varre essa caixa de entrada e cruza os recibos
# encontrados com o log de envio (por e-mail do destinatário + assunto),
# atualizando o status de leitura de cada linha.

def localizar_pasta_entrada_caixa_compartilhada(namespace, alvo):
    """Procura, entre as caixas de correio disponíveis no perfil do Outlook
    atual, a que corresponde ao e-mail/nome de exibição informado e retorna
    sua pasta "Caixa de Entrada". Retorna None se não encontrar (ex.: a
    caixa compartilhada não foi adicionada ao perfil)."""
    nomes_entrada = ["Caixa de Entrada", "Inbox"]
    alvo_lower = alvo.strip().lower()

    for pasta_raiz in namespace.Folders:
        try:
            nome_raiz = (pasta_raiz.Name or "").strip().lower()
        except Exception:
            continue
        if alvo_lower in nome_raiz:
            for nome_entrada in nomes_entrada:
                try:
                    return pasta_raiz.Folders(nome_entrada)
                except Exception:
                    continue
    return None


def extrair_assunto_original_do_recibo(assunto_recibo):
    """Remove prefixos comuns (pt/en) que o Outlook adiciona ao assunto dos
    relatórios de recibo, devolvendo o assunto original do e-mail enviado."""
    prefixos = ["lido:", "read:", "recusado:", "declined:", "entregue:", "delivered:"]
    texto = (assunto_recibo or "").strip()
    baixo = texto.lower()
    for prefixo in prefixos:
        if baixo.startswith(prefixo):
            return texto[len(prefixo):].strip()
    return texto


def perguntar_caixa_para_verificacao(namespace):
    """Reaproveita a mesma detecção de caixas do perfil do Outlook (contas
    completas e caixas compartilhadas/pastas adicionais) usada na etapa de
    seleção da caixa de envio, para que o usuário escolha de forma
    consistente qual caixa consultar em busca dos recibos de leitura."""
    caixas = listar_caixas_disponiveis(namespace)

    print("Caixas detectadas no perfil do Outlook:")
    for i, caixa in enumerate(caixas, start=1):
        rotulo_tipo = "conta completa" if caixa["tipo"] == "conta" else "pasta adicional / possível caixa compartilhada"
        print(f"  {i}. {caixa['rotulo']} ({rotulo_tipo})")

    while True:
        escolha = input(
            "E-mail ou nome de exibição da caixa usada no envio original "
            "(número da lista acima ou digite o e-mail/nome): "
        ).strip()
        if escolha.isdigit() and 1 <= int(escolha) <= len(caixas):
            return caixas[int(escolha) - 1]["rotulo"]
        if escolha:
            return escolha
        print("Valor vazio. Digite o número de uma opção da lista ou um e-mail/nome válido.")


def obter_email_smtp_do_recibo(item):
    """Para contatos internos do Exchange, SenderEmailAddress do item de
    recibo às vezes vem no formato X.500 (legacyExchangeDN) em vez do
    endereço SMTP, o que quebraria a comparação com a planilha/log. Esta
    função tenta algumas formas de obter o endereço SMTP real."""
    try:
        bruto = str(getattr(item, "SenderEmailAddress", "") or "").strip()
    except Exception:
        bruto = ""
    if "@" in bruto:
        return bruto.lower()

    try:
        smtp = item.PropertyAccessor.GetProperty(
            "http://schemas.microsoft.com/mapi/proptag/0x5D01001F"
        )
        if smtp and "@" in smtp:
            return smtp.strip().lower()
    except Exception:
        pass

    try:
        sender = item.Sender
        if sender is not None and sender.AddressEntry.Type == "EX":
            smtp = sender.AddressEntry.GetExchangeUser().PrimarySmtpAddress
            if smtp and "@" in smtp:
                return smtp.strip().lower()
    except Exception:
        pass

    return bruto.lower()


def verificar_confirmacoes_leitura(namespace):
    print("=" * 70)
    print("VERIFICAÇÃO DE CONFIRMAÇÕES DE LEITURA")
    print("=" * 70)
    import pandas as pd

    while True:
        caminho_log = limpar_caminho(input("Caminho do log de envio (log_envio.xlsx): "))
        if os.path.exists(caminho_log) and caminho_log.lower().endswith((".xlsx", ".xls")):
            break
        print("Arquivo não encontrado ou formato inválido (use .xlsx/.xls). Tente novamente.")

    try:
        df_log = pd.read_excel(caminho_log)
    except Exception as e:
        print(f"[ERRO] Não foi possível ler o log ({e}).")
        return

    colunas_esperadas = {"Nome", "Email", "Status"}
    if not colunas_esperadas.issubset(set(df_log.columns)):
        print(f"[ERRO] O log precisa conter as colunas {colunas_esperadas}. "
              f"Colunas encontradas: {list(df_log.columns)}")
        return

    if "Assunto" not in df_log.columns:
        df_log["Assunto"] = ""
    if "ConfirmacaoLeitura" not in df_log.columns:
        df_log["ConfirmacaoLeitura"] = "Pendente"
    if "DataHoraLeitura" not in df_log.columns:
        df_log["DataHoraLeitura"] = ""

    # Se o log já existia e essas colunas ficaram totalmente vazias/em branco
    # (ex.: gerado numa execução anterior deste script antes de haver
    # recibos), o pandas/Excel pode tê-las lido de volta como dtype
    # numérico (float64, coluna só de NaN). Forçamos para texto para poder
    # gravar valores como "Confirmada"/data e hora sem erro de tipo.
    df_log["Assunto"] = df_log["Assunto"].astype(object).where(df_log["Assunto"].notna(), "")
    df_log["ConfirmacaoLeitura"] = df_log["ConfirmacaoLeitura"].astype(object).where(
        df_log["ConfirmacaoLeitura"].notna(), "Pendente"
    )
    df_log["DataHoraLeitura"] = df_log["DataHoraLeitura"].astype(object).where(
        df_log["DataHoraLeitura"].notna(), ""
    )

    alvo = perguntar_caixa_para_verificacao(namespace)

    print(f"\nProcurando a Caixa de Entrada de '{alvo}' no perfil do Outlook...")
    pasta_entrada_alvo = localizar_pasta_entrada_caixa_compartilhada(namespace, alvo)
    pasta_entrada_padrao = namespace.GetDefaultFolder(6)  # olFolderInbox

    pastas_para_buscar = []
    if pasta_entrada_alvo is None:
        print(f"[AVISO] '{alvo}' não foi encontrada como mailbox adicional neste perfil.")
        print("        Verificando apenas na Caixa de Entrada padrão do usuário atual.")
    else:
        print("Caixa de entrada da conta de envio localizada.")
        pastas_para_buscar.append(pasta_entrada_alvo)

    # Quando o envio original foi feito "em nome de" uma caixa compartilhada
    # (SentOnBehalfOfName), o Exchange normalmente entrega os recibos de
    # leitura na Caixa de Entrada da conta pessoal que de fato enviou a
    # mensagem, e não na Caixa de Entrada própria da caixa compartilhada.
    # Por isso, sempre verificamos também a Caixa de Entrada padrão do
    # usuário atual, além da caixa escolhida (evitando duplicidade se forem
    # a mesma pasta).
    if pasta_entrada_alvo is None or pasta_entrada_alvo.EntryID != pasta_entrada_padrao.EntryID:
        pastas_para_buscar.append(pasta_entrada_padrao)

    # Em vez de usar Items.Restrict() com uma query DASL sobre o MessageClass
    # (que se mostrou pouco confiável dependendo da versão/configuração do
    # Outlook), percorremos todos os itens da pasta e checamos o
    # MessageClass diretamente em Python — mais lento, porém muito mais
    # confiável para encontrar os recibos de leitura/recusa.
    itens_recibo = []
    ids_vistos = set()
    candidatos_nao_reconhecidos = []
    prefixos_assunto_recibo = ("lido:", "read:", "recusado:", "declined:", "entregue:", "delivered:")
    for pasta in pastas_para_buscar:
        try:
            itens_pasta = pasta.Items
            total_itens = itens_pasta.Count
        except Exception as e:
            print(f"[ERRO] Falha ao acessar os itens de '{pasta.Name}': {e}")
            continue
        print(f"Verificando {total_itens} item(ns) em '{pasta.Name}'...")
        for item in itens_pasta:
            try:
                classe_item = str(item.MessageClass)
            except Exception:
                continue
            if not (classe_item.startswith("REPORT.IPM.Note.IPNRN") or classe_item.startswith("REPORT.IPM.Note.IPNNRN")
                    or classe_item.startswith("IPM.Note.IPNRN") or classe_item.startswith("IPM.Note.IPNNRN")):
                # Diagnóstico: se o assunto parece ser de um recibo (prefixo
                # "Lido:"/"Read:"/etc.) mas o MessageClass não bateu com o
                # esperado, guardamos para mostrar ao usuário — ajuda a
                # identificar rapidamente se o Outlook está usando uma classe
                # diferente da documentada.
                try:
                    assunto_item = str(item.Subject or "")
                except Exception:
                    assunto_item = ""
                if assunto_item.strip().lower().startswith(prefixos_assunto_recibo):
                    candidatos_nao_reconhecidos.append((assunto_item, classe_item))
                continue
            try:
                entry_id = item.EntryID
            except Exception:
                entry_id = None
            if entry_id is not None and entry_id in ids_vistos:
                continue
            if entry_id is not None:
                ids_vistos.add(entry_id)
            itens_recibo.append(item)

    if not itens_recibo:
        print("Nenhum recibo de leitura encontrado nas caixas verificadas.")
        if candidatos_nao_reconhecidos:
            print("\n[DIAGNÓSTICO] Encontrado(s) item(ns) com assunto de recibo, mas com")
            print("MessageClass diferente do esperado (IPM.Note.IPNRN/IPNNRN):")
            for assunto_item, classe_item in candidatos_nao_reconhecidos:
                print(f"  - Assunto: {assunto_item!r} | MessageClass: {classe_item!r}")

    # Em vez de recomputar (e reescanear) as colunas Email/Assunto do log
    # inteiro para cada recibo, construímos aqui um índice único: uma
    # varredura O(linhas_do_log) que permite, para cada recibo, uma busca
    # O(1) em vez de O(linhas_do_log) — importante quando o log é grande e
    # há muitos recibos para cruzar.
    from collections import defaultdict

    indice_por_email_e_assunto = defaultdict(list)
    indice_por_email_sem_assunto = defaultdict(list)
    email_normalizado = df_log["Email"].astype(str).str.strip().str.lower()
    assunto_normalizado = df_log["Assunto"].astype(str).str.strip().str.lower()
    status_normalizado = df_log["Status"].astype(str).str.strip()
    for idx in df_log.index[status_normalizado == "Enviado"]:
        email_idx = email_normalizado[idx]
        assunto_idx = assunto_normalizado[idx]
        indice_por_email_e_assunto[(email_idx, assunto_idx)].append(idx)
        if assunto_idx == "":
            indice_por_email_sem_assunto[email_idx].append(idx)

    total_confirmados = 0
    total_recusados = 0
    for item in itens_recibo:
        try:
            classe = str(item.MessageClass)
            email_leitor = obter_email_smtp_do_recibo(item)
            assunto_original = extrair_assunto_original_do_recibo(str(item.Subject or ""))
        except Exception as e:
            print(f"[DIAGNÓSTICO] Falha ao processar um recibo (item ignorado): {e}")
            continue

        # Itens de relatório (recibo de leitura/recusa) às vezes não expõem
        # ReceivedTime da mesma forma que um e-mail comum; usamos
        # CreationTime como alternativa e, no limite, deixamos em branco em
        # vez de descartar o recibo inteiro. O valor é formatado como
        # DD/MM/AAAA HH:MM para facilitar a leitura no log.
        try:
            data_leitura = item.ReceivedTime.strftime("%d/%m/%Y %H:%M")
        except Exception:
            try:
                data_leitura = item.CreationTime.strftime("%d/%m/%Y %H:%M")
            except Exception:
                data_leitura = ""

        recusado = classe.startswith("REPORT.IPM.Note.IPNNRN") or classe.startswith("IPM.Note.IPNNRN")

        indices_correspondentes = indice_por_email_e_assunto.get(
            (email_leitor, assunto_original.strip().lower()), []
        )
        if not indices_correspondentes:
            # Fallback: só entra em ação quando a própria linha do log não tem
            # assunto registrado (ex.: log gerado por uma execução antiga
            # deste script, antes de a coluna "Assunto" existir). Nesse caso,
            # casamos só pelo e-mail dentro deste log. NÃO usar esse fallback
            # quando a linha do log já tem um assunto diferente do recibo:
            # isso faria um recibo de um disparo anterior (ex.: "teste 7")
            # ser incorretamente atribuído a um disparo diferente para o
            # mesmo destinatário (ex.: "teste 9"), inflando as confirmações.
            indices_correspondentes = indice_por_email_sem_assunto.get(email_leitor, [])
        if not indices_correspondentes:
            print(
                f"[DIAGNÓSTICO] Recibo recebido de '{email_leitor}' (assunto do recibo: "
                f"'{assunto_original}') não encontrou nenhuma linha correspondente no log "
                f"(com Status='Enviado')."
            )
            continue

        status_novo = "Recusada" if recusado else "Confirmada"
        df_log.loc[indices_correspondentes, "ConfirmacaoLeitura"] = status_novo
        df_log.loc[indices_correspondentes, "DataHoraLeitura"] = data_leitura
        if recusado:
            total_recusados += 1
        else:
            total_confirmados += 1

    df_log.to_excel(caminho_log, index=False)

    n_enviados = int((df_log["Status"] == "Enviado").sum())
    n_pendentes = int(
        ((df_log["Status"] == "Enviado") & (df_log["ConfirmacaoLeitura"] == "Pendente")).sum()
    )
    print("\n" + "=" * 70)
    print("RESULTADO DA VERIFICAÇÃO")
    print("=" * 70)
    print(f"E-mails enviados no log:              {n_enviados}")
    print(f"Confirmações de leitura recebidas:    {total_confirmados}")
    print(f"Recusas de confirmação recebidas:     {total_recusados}")
    print(f"Ainda pendentes (sem resposta):       {n_pendentes}")
    print(f"\nLog atualizado salvo em: {caminho_log}")


# ==================== ORQUESTRAÇÃO ============================================

def escolher_modo_operacao():
    print("=" * 70)
    print("O que você deseja fazer?")
    print("=" * 70)
    print("  1. Novo disparo de e-mails")
    print("  2. Verificar confirmações de leitura de um disparo já feito")
    while True:
        escolha = input("Escolha uma opção (1/2): ").strip()
        if escolha in ("1", "2"):
            return escolha
        print("Opção inválida. Digite 1 ou 2.")


def main():
    modo = escolher_modo_operacao()
    print()

    executar_verificacao_ambiente()

    import win32com.client as win32
    outlook = win32.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")

    if modo == "2":
        verificar_confirmacoes_leitura(namespace)
        return

    caixas_disponiveis = listar_caixas_disponiveis(namespace)
    email_usuario = perguntar_email_usuario(caixas_disponiveis)
    caixa_envio = perguntar_caixa_envio(namespace, caixas_disponiveis)
    planilha = selecionar_planilha()
    _, assinatura_docx_html = perguntar_caminho_assinatura()
    logo_path = perguntar_logo(assinatura_docx_html)
    assinatura_html = montar_assinatura_final(assinatura_docx_html, logo_path)
    template_path, corpo_html_template = selecionar_template(assinatura_html)

    assunto = input("Assunto do e-mail: ").strip()
    solicitar_confirmacao_leitura = perguntar_confirmacao_leitura()
    gerar_log, caminho_log = perguntar_geracao_log(planilha)

    if not mostrar_resumo_e_confirmar(email_usuario, planilha, caixa_envio, logo_path, template_path, assunto,
                                       solicitar_confirmacao_leitura, gerar_log, caminho_log):
        print("Operação cancelada pelo usuário.")
        return

    enviar_emails(outlook, caixa_envio, planilha, logo_path, corpo_html_template, assunto,
                  solicitar_confirmacao_leitura, gerar_log, caminho_log)


if __name__ == "__main__":
    main()
