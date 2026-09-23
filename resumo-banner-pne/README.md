# Resumo — Banner do Simpósio de PNE (JOCAM / UNIP Marquês)

Arquivo para enviar à professora antes da submissão: **`Resumo_Banner_PNE_TEA_TDAH_TAG.docx`**

## Tema

Revisão de literatura sobre o manejo odontológico do paciente com **TEA nível 3 + TDAH + TAG**
(Opção 1 do enunciado: revisão sobre a condição da paciente atendida no box, com fotos no banner).

O resumo trata do manejo **de modo geral**, como a literatura recomenda que seja conduzido o
condicionamento, e não da conduta específica planejada para a paciente — que ainda está em
acompanhamento. O caso aparece no banner, pelas fotos e pela apresentação oral.

## Conformidade com as regras de submissão

| Regra | Situação |
|---|---|
| Até 260 palavras, **contando o título** | **251** (título 20 + corpo 231) |
| Introdução, objetivos, resultados e conclusão | na ordem, em parágrafo único |
| Materiais e métodos | **sem frase explícita**, por decisão da autora (ver observação) |
| Parágrafo único, sem tópicos (orientação da professora) | sim |
| Arial 12, justificado, entrelinhas 1,5 | sim |
| Título ARIAL 14 EM CAIXA ALTA | sim |
| Referências explícitas (não contam palavras) | 7 referências ABNT ao final |
| Até 5 participantes (1 autor + até 4 coautores) + 1 orientador | 1 autora + 1 coautora; **2 orientadores** (o formulário prevê 1 — confirmar) |
| 3 palavras-chave (orientação da professora) | sim |

## Participantes

- Autores: Thamyres Tosarelli e Beatriz Gouvea
- Orientadores: Rosemary Baptista Martins Teixeira e Ricardo Matsura Kodama

Paciente de 8 anos (nome não citado, por sigilo). Para as fotos do banner é
necessário termo de consentimento assinado pela responsável.

## Referências usadas

1. AAPD. *Behavior guidance for the pediatric dental patient*. Reference Manual of Pediatric Dentistry, 2024.
2. ALBHAISI et al. *BMC Oral Health*, 2022 — técnicas psicológicas no manejo de crianças com TEA.
3. BALIAN et al. *Int J Environ Res Public Health*, 2021 — pedagogia visual (metanálise).
4. CERMAK et al. *J Autism Dev Disord*, 2015 — ambiente sensorialmente adaptado (ECR piloto).
5. DRUMOND et al. *Caries Research*, 2022 — cárie em crianças com TDAH (metanálise).
6. STEIN DUKER et al. *JAMA Network Open*, 2023 — adaptação sensorial (ECR cruzado, 162 crianças).
7. BEZERRA; ASSIS; SANTOS. *Brazilian Journal of Health Review*, 2023 — atendimento odontológico de crianças com TEA.

## Como regerar o .docx

```bash
python3 gerar_docx.py   # requer python-docx; lê o texto de corpo.txt
```

---

## Versão alternativa: tema de tecnologia (Opção 2)

Mesmo pacote, com o tema do seminário do box — `Resumo_Banner_PNE_Tecnologia.docx`
e `Guia_de_estudo_Tecnologia.docx`.

| Regra | Situação |
|---|---|
| Até 270 palavras, sem contar título | **264 palavras** |
| Referências explícitas | 11 referências ABNT |
| Arial 12 justificado 1,5 / título Arial 14 caixa alta | sim |

Correção encontrada no `index.html` (slide de referências): a metanálise de Er:YAG está
atribuída ao *European Journal of Paediatric Dentistry*, 2018. O correto é **LI, Ting et al.
Lasers in Medical Science, v. 34, n. 2, p. 273-280, 2019**.


## Observação sobre materiais e métodos

Os dois resumos não trazem mais a frase "trata-se de uma revisão de literatura, com artigos
publicados entre 2015 e 2025" — o espaço foi usado para conteúdo clínico. O formulário lista
materiais e métodos entre os itens exigidos, então, se a orientadora cobrar, basta reinserir
uma destas frases logo após o objetivo e cortar o equivalente em palavras:

- TEA: "Para isso, foram analisados ensaios clínicos randomizados, revisões sistemáticas e
  diretrizes internacionais sobre manejo comportamental."
- Tecnologia: "Para isso, foram analisados ensaios clínicos randomizados, revisões sistemáticas
  e diretrizes nacionais e internacionais sobre cada recurso."

## PDFs

Os `.pdf` ao lado de cada `.docx` foram gerados por LibreOffice nesta máquina, que não tem
Arial instalada e substituiu por Liberation Sans — fonte metricamente idêntica, então a
quebra de linha e a paginação são as mesmas. Os `.docx` continuam especificando Arial.
Para um PDF em Arial de verdade, basta abrir o `.docx` no Word e salvar como PDF.

Regerar todos:

```bash
soffice --headless --convert-to pdf --outdir . *.docx
```
