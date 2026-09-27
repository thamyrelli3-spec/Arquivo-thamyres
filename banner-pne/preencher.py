# -*- coding: utf-8 -*-
"""Preenche o layout oficial do banner da V JOCAM mantendo a formatação do modelo."""
import re, html

SLIDE = 'unpacked/ppt/slides/slide1.xml'
CM = 360000  # EMU por centímetro

TITULO = "MANEJO ODONTOLÓGICO DO PACIENTE COM TEA NÍVEL 3, TDAH E TAG"

AUTORES = [
    "Autores*: Thamyres Tosarelli, Beatriz Gouvea",
    "Orientadores**: Rosemary Baptista Martins Teixeira, Ricardo Matsura Kodama",
]

INTRODUCAO = ("O Transtorno do Espectro Autista (TEA) nível 3 exige maior suporte, com comunicação verbal "
 "limitada, comportamentos repetitivos e alterações sensoriais. Com frequência vem acompanhado "
 "do Transtorno de Déficit de Atenção e Hiperatividade (TDAH) e do Transtorno de Ansiedade "
 "Generalizada (TAG), o que reduz a colaboração no consultório. Esses pacientes têm maior risco de cárie e de doença periodontal, "
 "pois a higiene bucal depende do cuidador, a dieta costuma ser seletiva e cariogênica e os "
 "psicofármacos reduzem o fluxo salivar.")

METODOS = ("O atendimento começa por anamnese detalhada com o responsável, que informa a rotina "
 "da criança, a forma de comunicação e os estímulos que causam desconforto. As consultas devem ser "
 "curtas, no mesmo horário, na mesma sala e com a mesma equipe, sendo a primeira apenas de "
 "adaptação, sem procedimento clínico. O condicionamento é gradual, por dessensibilização, "
 "associada ao dizer-mostrar-fazer, à pedagogia visual com pranchas e agendas de figuras, a vídeos "
 "que mostram o procedimento antes da consulta e ao reforço positivo a cada etapa cumprida. O "
 "ambiente é adaptado com redução de luz, ruído e odores, mantendo-se o cuidador presente.")

RESULTADOS = [
 ("A pedagogia visual aumenta a cooperação e melhora a qualidade da escovação, e a modelagem por "
  "vídeo reduz o número de consultas necessárias para o atendimento não invasivo. A adaptação "
  "sensorial do ambiente reduz o estresse fisiológico e comportamental em todas as fases da "
  "consulta."),
 ("A contenção protetora, a sedação e a anestesia geral permanecem indicadas apenas quando o "
  "manejo básico não é suficiente, mediante consentimento informado. Os estudos disponíveis, "
  "porém, têm amostras pequenas e grande heterogeneidade metodológica, o que impede a definição "
  "de um protocolo universal."),
]

CONCLUSAO = ("Não existe um protocolo único para esse perfil de paciente. O sucesso do atendimento "
 "depende da individualização, da previsibilidade, da capacitação da equipe e do vínculo com a "
 "família, e a ênfase na prevenção evita a necessidade de procedimentos mais invasivos.")

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
    novo_run = re.sub(r'<a:t>.*?</a:t>', '<a:t>%s</a:t>' % esc(texto), run.group(0), flags=re.S)
    return '<a:p>' + (pPr.group(0) if pPr else '') + novo_run + (end.group(0) if end else '') + '</a:p>'

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
shapes = re.findall(r'<p:sp>.*?</p:sp>', x, re.S)

def achar(trecho):
    for sp in shapes:
        if trecho in sp:
            return sp
    raise SystemExit('não achei shape com %r' % trecho)

# --- título -----------------------------------------------------------------
sp = achar('<a:t>TÍTULO</a:t>')
x = x.replace(sp, trocar_corpo(sp, [molde(paragrafos(sp)[0], TITULO)]))

# --- autores ----------------------------------------------------------------
sp = achar('Autores*:')
ps = paragrafos(sp)
x = x.replace(sp, trocar_corpo(sp, [molde(ps[0], AUTORES[0]), molde(ps[1], AUTORES[1])]))

# --- seções -----------------------------------------------------------------
secoes = [
    ('<a:t>Introdução/ Revisão de Literatura</a:t>', 'Introdução/ Revisão de Literatura', [INTRODUCAO]),
    ('<a:t>Material e Métodos</a:t>',                 'Material e Métodos',                [METODOS]),
    ('<a:t>Resultados</a:t>',                         'Resultados',                        RESULTADOS),
    ('<a:t>Conclusão</a:t>',                          'Conclusão',                         [CONCLUSAO]),
]
for marca, titulo_sec, blocos in secoes:
    sp = achar(marca)
    ps = paragrafos(sp)
    cab = por_tamanho(ps, 6600)
    def corpo_par(texto):
        t = molde(cab, texto)
        t = t.replace('sz="6600"', 'sz="4800"').replace('val="6600"', 'val="4800"')
        return re.sub(r'(<a:rPr )b="1"', r'\1b="0"', t)
    novos = [molde(cab, titulo_sec)] + [corpo_par(b) for b in blocos]
    x = x.replace(sp, trocar_corpo(sp, novos))

# --- referências ------------------------------------------------------------
sp = achar('<a:t>Referências</a:t>')
ps = paragrafos(sp)
cab, corpo = por_tamanho(ps, 2600), por_tamanho(ps, 1800)
novos = ([molde(cab, 'Referências:')]
         + [molde(corpo, r) for r in REFERENCIAS]
         + [molde(corpo, r) for r in RODAPE]
         + [molde(cab, 'Palavras-chave:'), molde(corpo, PALAVRAS)])
x = x.replace(sp, trocar_corpo(sp, novos))

# --- reposiciona blocos ------------------------------------------------------
def mover(marca, y_cm):
    global x
    m = re.search(r'<p:sp>(?:(?!</p:sp>).)*?' + re.escape(marca) + r'.*?</p:sp>', x, re.S)
    bloco = m.group(0)
    off = re.search(r'<a:off x="(-?\d+)" y="(-?\d+)"/>', bloco)
    novo = '<a:off x="%s" y="%d"/>' % (off.group(1), int(round(y_cm * CM)))
    x = x.replace(bloco, bloco.replace(off.group(0), novo, 1), 1)
    print('   %-22s y: %.1f -> %.1f cm' % (marca[:22], int(off.group(2)) / CM, y_cm))

mover(esc(TITULO)[:30], 17.8)          # abaixo do logo da JOCAM
mover('Autores*:', 24.2)
mover('Introdução/ Revisão', 28.2)
mover('<a:t>Material e Métodos</a:t>', 43.2)  # folga para a introdução
mover('Referências:', 105.5)           # cabia fora da folha

open(SLIDE, 'w', encoding='utf-8').write(x)
print('slide preenchido — %d caracteres' % len(x))
