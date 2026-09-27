# Banner — III Simpósio de PNE (V JOCAM)

`Banner_PNE_TEA.pptx` — layout oficial preenchido. `Banner_PNE_TEA.pdf` é o mesmo para a gráfica.

- Layout: **90 × 120 cm**, tudo em Arial
- Título 80pt negrito · seções 66pt negrito · corpo 48pt · legendas 24pt · referências 18pt

## Estrutura

Segue o formato dos banners aprovados em edições anteriores, e não a lista completa do modelo:

**Introdução → Revisão de Literatura → três figuras com legenda → Conclusão → Referências**

As caixas de *Material e Métodos* e *Resultados* do modelo não fazem sentido numa revisão de
literatura: a primeira foi reaproveitada como Revisão de Literatura e a segunda foi removida,
abrindo espaço para as figuras.

## O que foi mantido do modelo

Fonte, tamanhos, negrito, cores, fundo e logos. O preenchimento clona os parágrafos do próprio
modelo e troca só o texto, então nenhuma formatação foi redesenhada.

## O que precisou ser deslocado

O modelo prevê título de uma linha. Com o título real em 80pt o texto passava por cima do logo
da JOCAM. As caixas desceram ou subiram apenas o necessário para não haver sobreposição:

| Caixa | De | Para |
|---|---|---|
| Título | 12,4 cm | 17,8 cm |
| Autores/Orientadores | 16,7 cm | 24,2 cm |
| Introdução | 23,9 cm | 28,2 cm |
| Revisão de Literatura | 41,2 cm | 40,0 cm |
| Conclusão | 96,3 cm | 93,0 cm |
| Referências | 111,0 cm | 104,0 cm |

## Fotos

Três molduras tracejadas, de 26,5 × 21 cm, a partir de 68 cm. Para inserir sem bagunçar o
alinhamento: clique na moldura, **Formatar Forma → Preenchimento → Imagem**. A foto entra no
lugar exato da moldura. Depois apague o texto "INSIRA A FOTO N" e ajuste a legenda.

As legendas atuais são sugestões — troque pelo que a foto mostra de fato.

Foto de paciente exige **termo de consentimento assinado pela responsável**.

## Regerar

```bash
python3 -c "import zipfile;zipfile.ZipFile('modelo_banner.pptx').extractall('unpacked')"
python3 preencher.py
rm -f Banner_PNE_TEA.pptx && (cd unpacked && zip -Xrq ../Banner_PNE_TEA.pptx .)
```
