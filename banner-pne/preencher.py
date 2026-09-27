# -*- coding: utf-8 -*-
"""Preenche o layout oficial do banner da V JOCAM mantendo a formatação do modelo.

Estrutura adotada, seguindo os banners aprovados em edições anteriores:
Introdução -> Revisão de Literatura -> três figuras -> Conclusão -> Referências.
"""
import re, html

SLIDE = 'unpacked/ppt/slides/slide1.xml'
CM = 360000  # EMU por centímetro

TITULO = "MANEJO ODONTOLÓGICO DO PACIENTE COM TEA NÍVEL 3, TDAH E TAG"

AUTORES = [
    "Autores*: Tosarelli, T.; Gouvea, B.",
    "Orientadores**: Teixeira, R. B. M.; Kodama, R. M.",
]

INTRODUCAO = ("O Transtorno do Espectro Autista (TEA) é um distúrbio do neurodesenvolvimento que "
 "altera a comunicação, a interação social e o comportamento. No nível 3, o mais grave, o paciente "
 "necessita de apoio muito substancial: a comunicação verbal é limitada, há comportamentos "
 "repetitivos e alterações no processamento sensorial, o que torna o atendimento odontológico um "
 "desafio para toda a equipe.")

REVISAO = [
 ("A associação do TEA ao Transtorno de Déficit de Atenção e Hiperatividade (TDAH) e ao Transtorno "
  "de Ansiedade Generalizada (TAG) é frequente e reduz ainda mais a colaboração na cadeira "
  "odontológica. Esses pacientes apresentam maior risco de cárie e de doença periodontal, pois a "
  "higiene bucal depende do cuidador, a dieta costuma ser seletiva e cariogênica e os psicofármacos "
  "reduzem o fluxo salivar."),
 ("A colaboração é construída, e não obtida de imediato. A anamnese deve ser detalhada com o "
  "responsável, identificando a rotina, a comunicação e os estímulos que causam desconforto. As "
  "consultas devem ser curtas, no mesmo horário, sala e equipe, sendo a primeira apenas de "
  "adaptação. A dessensibilização, o dizer-mostrar-fazer, a "
  "pedagogia visual com pranchas e agendas de figuras, a modelagem por vídeo e o reforço positivo "
  "aumentam a cooperação e reduzem o número de consultas necessárias. A adaptação sensorial do "
  "ambiente, com redução de luz, ruído e odores, diminui o estresse fisiológico e comportamental. "
  "Contenção protetora, sedação e anestesia geral ficam reservadas à falha do manejo básico."),
]

LEGENDAS = [
 "Fig. 1 - Condicionamento da paciente antes do atendimento",
 "Fig. 2 - Pedagogia visual durante a consulta",
 "Fig. 3 - Adequação do meio bucal",
]

CONCLUSAO = ("Não existe protocolo único para o paciente com TEA. O sucesso do atendimento depende "
 "da individualização, da previsibilidade, da capacitação da equipe e do vínculo com a família. A "
 "ênfase na prevenção reduz a necessidade de contenção, sedação e anestesia geral.")

REFERENCIAS = [
 "ALBHAISI, I. N. et al. Effectiveness of psychological techniques in dental management for children with autism spectrum disorder: a systematic literature review. BMC Oral Health, v. 22, 2022.",
 "AMERICAN ACADEMY OF PEDIATRIC DENTISTRY. Behavior guidance for the pediatric dental patient. In: The Reference Manual of Pediatric Dentistry. Chicago: AAPD, 2024. p. 358-378.",
 "BALIAN, A. et al. Is visual pedagogy effective in improving cooperation towards oral hygiene and dental care in children with autism spectrum disorder? Int. J. Environ. Res. Public Health, v. 18, n. 2, p. 789, 2021.",
 "BEZERRA, R. C.; ASSIS, J. A.; SANTOS, P. U. O atendimento odontológico à crianças com Transtorno do Espectro Autista: uma revisão de literatura. Brazilian Journal of Health Review, v. 6, n. 3, p. 13155-13171, 2023.",
 "CERMAK, S. A. et al. Sensory adapted dental environments to enhance oral care for children with autism spectrum disorders. J. Autism Dev. Disord., v. 45, n. 9, p. 2876-2888, 2015.",
 "DA SILVA MORO, J. et al. Efficacy of the video modeling technique as a facilitator of non-invasive dental care in autistic children: randomized clinical trial. J. Autism Dev. Disord., v. 54, n. 2, p. 501-508, 2024.",
 "DRUMOND, V. Z. et al. Dental caries in children with attention deficit/hyperactivity disorder: a meta-analysis. Caries Research, v. 56, n. 1, p. 3-14, 2022.",
 "STEIN DUKER, L. I. et al. Sensory adaptations to improve physiological and behavioral distress during dental visits in autistic children. JAMA Network Open, v. 6, n. 6, e2316346, 2023.",
]
RODAPE = [
 "*Thamyres Tosarelli, Beatriz Gouvea - Campus Marquês, São Paulo.",
 "**Rosemary Baptista Martins Teixeira, Ricardo Matsura Kodama - UNIP",
]
PALAVRAS = "Transtorno do Espectro Autista; Assistência Odontológica para Pessoas com Deficiências; Manejo Comportamental."

# --- posições finais, em cm -------------------------------------------------
POS = {
    'titulo':    17.8,   # abaixo do logo da JOCAM
    'autores':   24.2,
    'introducao':28.2,
    'revisao':   40.0,
    'conclusao': 93.0,
    'referencias': 104.0,
}
FOTO_Y, FOTO_H, FOTO_W = 68.0, 21.0, 26.5
FOTO_X = [2.0, 31.75, 61.5]
LEGENDA_Y = 89.5

# ----------------------------------------------------------------------------

def esc(t):
    return html.escape(t, quote=False)

def paragrafos(xml):
    return re.findall(r'<a:p>.*?</a:p>', xml, re.S)

def molde(par, texto):
    """Clona um parágrafo do modelo trocando só o texto, preservando pPr e rPr."""
    pPr = re.search(r'<a:pPr.*?</a:pPr>', par, re.S)
    run = re.search(r'<a:r>.*?</a:r>', par, re.S)
    end = re.search(r'<a:endParaRPr.*?</a:endParaRPr>', par, re.S)
    novo = re.sub(r'<a:t>.*?</a:t>', '<a:t>%s</a:t>' % esc(texto), run.group(0), flags=re.S)
    return '<a:p>' + (pPr.group(0) if pPr else '') + novo + (end.group(0) if end else '') + '</a:p>'

def trocar_corpo(sp, paras):
    bodyPr = re.search(r'<a:bodyPr[^>]*>.*?</a:bodyPr>|<a:bodyPr[^>]*/>', sp, re.S).group(0)
    corpo = '<p:txBody>' + bodyPr + '<a:lstStyle/>' + ''.join(paras) + '</p:txBody>'
    return re.sub(r'<p:txBody>.*?</p:txBody>', lambda m: corpo, sp, flags=re.S)

def por_tamanho(pars, sz):
    for p in pars:
        if re.search(r'<a:rPr[^>]*sz="%d"' % sz, p):
            return p
    raise SystemExit('não achei parágrafo sz=%d' % sz)

x = open(SLIDE, encoding='utf-8').read()

def achar(trecho):
    for sp in re.findall(r'<p:sp>.*?</p:sp>', x, re.S):
        if trecho in sp:
            return sp
    raise SystemExit('não achei shape com %r' % trecho)

def preencher_secao(marca, titulo_sec, blocos):
    global x
    sp = achar(marca)
    ps = paragrafos(sp)
    cab = por_tamanho(ps, 6600)
    def corpo(texto):
        t = molde(cab, texto).replace('sz="6600"', 'sz="4800"').replace('val="6600"', 'val="4800"')
        return re.sub(r'(<a:rPr )b="1"', r'\1b="0"', t)
    x = x.replace(sp, trocar_corpo(sp, [molde(cab, titulo_sec)] + [corpo(b) for b in blocos]))

def mover(marca, y_cm):
    global x
    m = re.search(r'<p:sp>(?:(?!</p:sp>).)*?' + re.escape(marca) + r'.*?</p:sp>', x, re.S)
    bloco = m.group(0)
    off = re.search(r'<a:off x="(-?\d+)" y="(-?\d+)"/>', bloco)
    novo = '<a:off x="%s" y="%d"/>' % (off.group(1), int(round(y_cm * CM)))
    x = x.replace(bloco, bloco.replace(off.group(0), novo, 1), 1)

def apagar(marca):
    global x
    m = re.search(r'<p:sp>(?:(?!</p:sp>).)*?' + re.escape(marca) + r'.*?</p:sp>', x, re.S)
    x = x.replace(m.group(0), '', 1)

# --- título, autores --------------------------------------------------------
sp = achar('<a:t>TÍTULO</a:t>')
x = x.replace(sp, trocar_corpo(sp, [molde(paragrafos(sp)[0], TITULO)]))

sp = achar('Autores*:')
ps = paragrafos(sp)
x = x.replace(sp, trocar_corpo(sp, [molde(ps[0], AUTORES[0]), molde(ps[1], AUTORES[1])]))

# --- seções: a caixa de Material e Métodos vira Revisão de Literatura,
#     a de Resultados sai (o banner usa figuras nesse espaço) ----------------
preencher_secao('<a:t>Introdução/ Revisão de Literatura</a:t>', 'Introdução', [INTRODUCAO])
preencher_secao('<a:t>Material e Métodos</a:t>', 'Revisão de Literatura', REVISAO)
apagar('<a:t>Resultados</a:t>')
preencher_secao('<a:t>Conclusão</a:t>', 'Conclusão', [CONCLUSAO])

# --- referências ------------------------------------------------------------
sp = achar('<a:t>Referências</a:t>')
ps = paragrafos(sp)
cab, corpo = por_tamanho(ps, 2600), por_tamanho(ps, 1800)
x = x.replace(sp, trocar_corpo(sp,
      [molde(cab, 'Referências:')]
    + [molde(corpo, r) for r in REFERENCIAS]
    + [molde(corpo, r) for r in RODAPE]
    + [molde(cab, 'Palavras-chave:'), molde(corpo, PALAVRAS)]))

# --- molduras das fotos e legendas ------------------------------------------
ids = [int(i) for i in re.findall(r'<p:cNvPr id="(\d+)"', x)]
proximo = max(ids) + 1

def caixa(id_, nome, x_cm, y_cm, w_cm, h_cm, texto, sz, moldura):
    geo = ('<a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
           '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
           % (x_cm*CM, y_cm*CM, w_cm*CM, h_cm*CM))
    if moldura:
        geo += ('<a:solidFill><a:srgbClr val="FFFFFF"><a:alpha val="55000"/></a:srgbClr></a:solidFill>'
                '<a:ln w="28575"><a:solidFill><a:srgbClr val="000066"/></a:solidFill>'
                '<a:prstDash val="dash"/></a:ln>')
    else:
        geo += '<a:noFill/>'
    anchor = 'ctr' if moldura else 't'
    return ('<p:sp><p:nvSpPr><p:cNvPr id="%d" name="%s"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
            '<p:spPr>%s</p:spPr>'
            '<p:txBody><a:bodyPr anchor="%s" wrap="square" lIns="91425" rIns="91425" '
            'tIns="45700" bIns="45700"/><a:lstStyle/>'
            '<a:p><a:pPr algn="ctr"><a:buNone/></a:pPr>'
            '<a:r><a:rPr lang="pt-BR" sz="%d" b="0" i="0"><a:solidFill>'
            '<a:srgbClr val="000066"/></a:solidFill><a:latin typeface="Arial"/>'
            '<a:ea typeface="Arial"/><a:cs typeface="Arial"/></a:rPr><a:t>%s</a:t></a:r>'
            '</a:p></p:txBody></p:sp>'
            % (id_, nome, geo, anchor, sz, esc(texto)))

novas = []
for i, xf in enumerate(FOTO_X):
    novas.append(caixa(proximo + i, 'Moldura foto %d' % (i+1), xf, FOTO_Y, FOTO_W, FOTO_H,
                       'INSIRA A FOTO %d' % (i+1), 3600, True))
for i, xf in enumerate(FOTO_X):
    novas.append(caixa(proximo + 3 + i, 'Legenda %d' % (i+1), xf, LEGENDA_Y, FOTO_W, 2.4,
                       LEGENDAS[i], 2400, False))
x = x.replace('</p:spTree>', ''.join(novas) + '</p:spTree>')

# --- reposicionamento --------------------------------------------------------
mover(esc(TITULO)[:30], POS['titulo'])
mover('Autores*:', POS['autores'])
mover('<a:t>Introdução</a:t>', POS['introducao'])
mover('<a:t>Revisão de Literatura</a:t>', POS['revisao'])
mover('<a:t>Conclusão</a:t>', POS['conclusao'])
mover('Referências:', POS['referencias'])

open(SLIDE, 'w', encoding='utf-8').write(x)
print('slide montado — %d shapes' % len(re.findall(r'<p:sp>', x)))
