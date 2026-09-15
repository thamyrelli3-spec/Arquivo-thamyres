# -*- coding: utf-8 -*-
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

doc = Document()
s = doc.sections[0]
for a in ('top_margin','bottom_margin','left_margin','right_margin'):
    setattr(s, a, Cm(2.0))

def p(t='', tam=11, neg=False, ital=False, al=WD_ALIGN_PARAGRAPH.JUSTIFY, esp=1.15, dep=6, recuo=0):
    par = doc.add_paragraph(); par.alignment = al
    pf = par.paragraph_format; pf.line_spacing = esp
    pf.space_before = Pt(0); pf.space_after = Pt(dep)
    if recuo: pf.left_indent = Cm(recuo)
    r = par.add_run(t); r.bold = neg; r.italic = ital
    r.font.name = 'Arial'; r.font.size = Pt(tam)
    r._element.rPr.rFonts.set(qn('w:cs'), 'Arial')
    r._element.rPr.rFonts.set(qn('w:hAnsi'), 'Arial')
    return par

def h(t, tam=13): p(t, tam, True, False, WD_ALIGN_PARAGRAPH.LEFT, 1.15, 8)
def marca(t): p(t, 10, True, True, WD_ALIGN_PARAGRAPH.LEFT, 1.0, 2)
def fala(t): p(t, 11.5, False, False, WD_ALIGN_PARAGRAPH.JUSTIFY, 1.3, 10)
def b(t, tam=11): p('•  ' + t, tam, False, False, WD_ALIGN_PARAGRAPH.JUSTIFY, 1.15, 4, recuo=0.6)

p('FALA DE APRESENTAÇÃO DO BANNER', 14, True, False, WD_ALIGN_PARAGRAPH.CENTER, 1.15, 4)
p('Manejo odontológico do paciente com TEA nível 3, TDAH e TAG', 12, False, False,
  WD_ALIGN_PARAGRAPH.CENTER, 1.15, 4)
p('Simpósio de Pacientes com Necessidades Especiais', 11, False, True,
  WD_ALIGN_PARAGRAPH.CENTER, 1.15, 16)

h('VERSÃO COMPLETA — cerca de 5 minutos')
p('As marcações entre colchetes são só para você, não fazem parte da fala.', 10, False, True,
  WD_ALIGN_PARAGRAPH.LEFT, 1.15, 10)

marca('[ABERTURA — 20 segundos. Olhe para a pessoa, não para o banner.]')
fala('Bom dia! Meu nome é Thamyres Tosarelli, esse trabalho eu fiz junto com a Beatriz Gouvea, '
     'com orientação da professora Rosemary Martins. O nosso tema é o manejo odontológico do '
     'paciente com Transtorno do Espectro Autista nível 3, associado a TDAH e a Transtorno de '
     'Ansiedade Generalizada. Posso começar?')

marca('[POR QUE ESSE TEMA — 30 segundos. Aponte a foto da paciente.]')
fala('A gente escolheu esse tema por causa de uma paciente que a gente atende aqui na clínica de '
     'pacientes com necessidades especiais. Ela tem oito anos e tem esses três diagnósticos '
     'juntos. É uma paciente pouco colaborativa, e a partir da dificuldade que a gente encontrou '
     'no atendimento dela a gente foi buscar na literatura o que é recomendado para esse perfil.')

marca('[O QUE É CADA CONDIÇÃO — 1 minuto. Aponte o quadro das condições.]')
fala('O autismo é classificado em três níveis, de acordo com o quanto a pessoa precisa de apoio. '
     'O nível 3 é o mais grave, é o que a literatura chama de apoio muito substancial: a '
     'comunicação verbal é bem limitada, existem comportamentos repetitivos e, o que mais importa '
     'para a gente na odontologia, existe uma alteração no processamento sensorial. Ou seja, luz, '
     'ruído, cheiro e toque são sentidos de um jeito muito mais intenso do que a gente sente.')
fala('Junto com isso, essa paciente tem TDAH, que traz desatenção e impulsividade, então ela não '
     'consegue se manter parada na cadeira por muito tempo. E tem o transtorno de ansiedade '
     'generalizada, que faz com que ela já chegue no consultório em estado de alerta, antecipando '
     'que alguma coisa ruim vai acontecer. Quando essas três condições se somam, a colaboração cai '
     'muito.')

marca('[POR QUE ISSO IMPORTA NA ODONTOLOGIA — 45 segundos. Aponte a parte do risco de cárie.]')
fala('E isso tem consequência direta na saúde bucal. Esses pacientes têm mais cárie e mais doença '
     'periodontal, e não é pela condição em si. É pela soma de quatro coisas: a higiene depende do '
     'cuidador, então nem sempre é feita do jeito ideal; a dieta costuma ser seletiva, muitas vezes '
     'pastosa e açucarada, porque a criança recusa textura; a hipersensibilidade na boca faz a '
     'criança não aceitar a escova; e os medicamentos que ela usa, como metilfenidato e '
     'risperidona, reduzem o fluxo salivar. Uma metanálise de 2022 mostrou que crianças com TDAH '
     'têm mais de três vezes mais chance de ter cárie do que crianças sem o transtorno.')

marca('[O MIOLO: COMO FAZER O CONDICIONAMENTO — 1 minuto e 30. Vá acompanhando com a mão.]')
fala('Aí vem a pergunta principal do trabalho: como deixar esse paciente mais colaborativo? A '
     'literatura mostra que isso se constrói, não acontece de uma vez.')
fala('Primeiro, a anamnese tem que ser mais detalhada do que a habitual. A gente precisa perguntar '
     'ao responsável qual é a rotina da criança, como ela se comunica, se ela usa alguma prancha ou '
     'aplicativo, e quais são os estímulos que incomodam. Isso muda completamente a forma de '
     'conduzir a consulta.')
fala('Segundo, previsibilidade. As consultas devem ser curtas, sempre no mesmo horário, na mesma '
     'sala e com a mesma equipe. Para essa criança, mudar a sala ou o operador já é motivo de '
     'recusa. E a primeira consulta deve ser só de adaptação, sem nenhum procedimento. Parece '
     'perda de tempo, mas é o que faz a segunda consulta acontecer.')
fala('Terceiro, o condicionamento propriamente dito, que é gradual. A gente usa a dessensibilização, '
     'apresentando o consultório aos poucos, junto com o dizer-mostrar-fazer, com pranchas e agendas '
     'de figuras que mostram a sequência do que vai acontecer, com vídeos que a criança assiste em '
     'casa antes de vir, e com reforço positivo a cada etapa que ela consegue cumprir. A metanálise '
     'que a gente cita mostrou que a pedagogia visual melhora tanto a cooperação quanto a qualidade '
     'da escovação.')
fala('E quarto, o ambiente. Diminuir a luz, o barulho e o cheiro, e deixar o cuidador junto. Um '
     'ensaio clínico de 2023, com 162 crianças autistas, mostrou que só adaptar o ambiente já reduz '
     'o estresse fisiológico da criança, medido no corpo dela, em todas as fases da consulta.')

marca('[QUANDO ESCALAR — 30 segundos.]')
fala('Agora, isso não funciona com todo mundo. Quando o manejo básico não é suficiente, aí sim '
     'entram a contenção protetora, a sedação e a anestesia geral. Mas a diretriz da Academia '
     'Americana de Odontopediatria é clara: isso é manejo avançado, vem depois, com consentimento '
     'informado e registro em prontuário, e sempre na técnica menos restritiva possível.')

marca('[FECHAMENTO — 30 segundos. Volte a olhar para a pessoa.]')
fala('Para concluir: a nossa principal conclusão é que não existe um protocolo único. O autismo é '
     'muito heterogêneo e os próprios estudos reconhecem essa limitação. O que funciona é '
     'individualizar, construir junto com a família e insistir na prevenção, porque quanto mais a '
     'gente previne, menos a gente precisa recorrer à anestesia geral. Obrigada!')

doc.add_page_break()

h('VERSÃO CURTA — cerca de 2 minutos')
p('Para quando a pessoa disser "me resume rapidinho" ou quando a fila estiver grande.', 10, False,
  True, WD_ALIGN_PARAGRAPH.LEFT, 1.15, 10)
fala('Bom dia! Eu sou a Thamyres, esse trabalho é meu e da Beatriz, orientado pela professora '
     'Rosemary. A gente fez uma revisão de literatura sobre o manejo do paciente com autismo nível '
     '3 associado a TDAH e ansiedade generalizada, a partir de uma paciente de oito anos que a '
     'gente atende na clínica.')
fala('O ponto central é que esse paciente tem mais cárie e mais doença periodontal, mas não por '
     'causa do transtorno em si: é a higiene que depende do cuidador, a dieta seletiva, a '
     'hipersensibilidade na boca e a boca seca causada pela medicação.')
fala('E a literatura mostra que a colaboração é construída. Anamnese detalhada com o responsável, '
     'consultas curtas e previsíveis, sempre no mesmo horário e com a mesma equipe, primeira '
     'consulta só de adaptação, e o condicionamento gradual com dessensibilização, dizer-mostrar-'
     'fazer, pranchas de figuras, vídeo antes da consulta e reforço positivo. Adaptar o ambiente, '
     'diminuindo luz e ruído, reduz o estresse fisiológico da criança.')
fala('Contenção, sedação e anestesia geral só quando isso tudo não basta. E a conclusão é que não '
     'existe protocolo único: tem que individualizar e apostar na prevenção. Obrigada!')

doc.add_page_break()

h('SE TRAVAR OU ESQUECER')
for t in ['Se esquecer a sequência, use o banner como roteiro: leia o título do tópico e desenvolva '
          'a partir dele. Ninguém percebe.',
          'Se esquecer um número, não invente. Diga "a metanálise mostrou um risco significativamente '
          'maior" e siga.',
          'Se te interromperem no meio, responda a pergunta e depois retome com "voltando ao que eu '
          'estava falando...".',
          'Se perguntarem algo que você não sabe, diga que não sabe, mas diga o que sabe perto '
          'disso. "Esse dado específico eu não tenho, mas o que eu vi na literatura foi...".']:
    b(t)

p('', 8, dep=10)
h('DICAS DE POSTURA')
for t in ['Não leia o banner. Quem avalia já sabe ler. Você está ali para contar o que o banner não '
          'diz.',
          'Fique ao lado do banner, nunca na frente. Aponte com a mão aberta, não com o dedo.',
          'Fale com quem está perguntando, não com o painel.',
          'Comece perguntando "posso começar?" ou "quanto tempo eu tenho?". Mostra segurança e evita '
          'você ser cortada no meio.',
          'Se tiver foto da paciente no banner, comece por ela. Caso real prende a atenção mais do '
          'que qualquer introdução teórica.',
          'Treine em voz alta pelo menos três vezes, cronometrando. Ler mentalmente engana, você '
          'sempre demora mais falando.']:
    b(t)

p('', 8, dep=10)
h('OBSERVAÇÃO SOBRE UM DADO DA FALA')
p('Na versão completa há a possibilidade de citar que a modelagem por vídeo reduz o número de '
  'consultas necessárias, que vem de um ensaio clínico brasileiro (DA SILVA MORO et al., Journal '
  'of Autism and Developmental Disorders, v. 54, n. 2, p. 501-508, 2024). Esse estudo não está na '
  'lista de referências do resumo de TEA. Se quiser usar esse dado na apresentação, acrescente a '
  'referência na lista, para não citar na fala um trabalho que não aparece no banner.', 11)

doc.save('Fala_Apresentacao_Banner_TEA.docx')
print('ok')
