# Banner — III Simpósio de PNE (V JOCAM)

`Banner_PNE_TEA.pptx` — layout oficial preenchido. `Banner_PNE_TEA.pdf` é o mesmo para a gráfica.

- Tamanho do layout: **90 × 120 cm**
- Tudo em Arial, como no modelo: título 80pt negrito, seções 66pt negrito, corpo 48pt, referências 18pt

## O que foi mantido do modelo

Fonte, tamanhos, negrito, cores, fundo, logos e a ordem das seções. Nenhum elemento foi
redesenhado: o preenchimento clona os parágrafos do próprio modelo e troca só o texto.

## O que precisou ser deslocado (e por quê)

O modelo prevê um título de uma linha. Com o título real, em 80pt, o texto passava por cima
do logo da JOCAM e da linha de autores. Sem mudar nenhuma fonte ou tamanho, quatro caixas
desceram ou subiram para não haver sobreposição:

| Caixa | De | Para | Motivo |
|---|---|---|---|
| Título | 12,4 cm | 17,8 cm | ficava sobre o logo da JOCAM |
| Autores/Orientadores | 16,7 cm | 24,2 cm | acompanha o título de duas linhas |
| Introdução | 23,9 cm | 28,2 cm | acompanha os autores |
| Material e Métodos | 41,2 cm | 43,2 cm | a última linha da introdução encostava |
| Referências | 111,0 cm | 105,5 cm | as palavras-chave caíam fora da folha |

## Onde entram as fotos

A faixa livre entre **Resultados** e **Conclusão** (de ~83 cm a ~96 cm, largura toda) comporta
duas ou três fotos lado a lado. Há uma segunda faixa menor entre Material e Métodos e
Resultados (~61 a 67 cm).

Foto de paciente exige **termo de consentimento assinado pela responsável**.

## Regerar

```bash
python3 -c "import zipfile;zipfile.ZipFile('modelo_banner.pptx').extractall('unpacked')"
python3 preencher.py
rm -f Banner_PNE_TEA.pptx && (cd unpacked && zip -Xrq ../Banner_PNE_TEA.pptx .)
```
