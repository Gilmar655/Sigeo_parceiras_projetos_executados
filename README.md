# SIGEO Parceiras — Programação e localização de projetos

Painel web (estático, GitHub Pages) com as programações SIGEO das parceiras **Start**, **Conecta** e **Engelmig** — Enel Brasil Grids · Gestão de Programação.

## Funcionalidades
- **Painel geral**: indicadores (executado, parcial, cancelado, em aberto, efetividade, CHI, georreferenciamento) e gráficos interativos — clique em um gráfico para filtrar.
- **Mapa de projetos**: localização de cada projeto (coordenada do equipamento mais programado — base Buscapé), agrupamento automático, cor por último status ou por parceira, mapa claro/OSM/satélite, links Google Maps e Waze.
- **Histórico do projeto**: ao clicar em qualquer projeto (tabela, mapa ou programação) abre a linha do tempo com **todas as programações**, cadastro do projeto, mapa dos equipamentos, gráfico de status no tempo e exportação Excel. Link direto: `#projeto=CODIGO`.
- **Projetos / Programações**: tabelas com ordenação, paginação e pesquisa.
- **Parceiras**: comparativo, efetividade e volume mensal, mix de intervenções, cancelamentos.
- **Filtros**: pesquisa livre, período, parceira, status, tipo, região, contratada, descrição, status do projeto e localização.
- **Exportar**: Excel do recorte (abas Programações, Projetos, Resumo, Por parceira/mês/contratada/tipo, com autofiltro), Excel completo, projetos com coordenadas, CSV, modelo de importação e `base.js`.
- **Importar**: arraste planilhas Excel (programações, cadastro de projetos, equipamentos). Reconhecimento automático das colunas; base salva no navegador (IndexedDB).

## Publicar
1. Envie todo o conteúdo desta pasta para a raiz de um repositório novo.
2. *Settings → Pages → Deploy from a branch → main / (root)*.

## Atualizar a base publicada
Importe as planilhas novas em **Importar / Exportar**, clique em **base.js para publicar no GitHub** e substitua `data/base.js` no repositório.

## Base atual
Programações 01/01/2026 a 16/11/2026: Start 1.992 · Conecta 13.047 · Engelmig 669 (15.708 linhas, 6.591 projetos, 6.527 georreferenciados).
Fontes: Start_16_11_2026.xlsx, Conecta_16_11_2026.xlsx, Engelmig_16_11.xlsx, Projetos_01_01_31_maio.xlsx, Equipamentos_Buscape_Links.xlsx (base 12/06/2026).

## Bibliotecas (incluídas em `vendor/`, sem CDN)
Chart.js 4.4.4 · Leaflet 1.9.4 + MarkerCluster 1.5.3 · SheetJS 0.18.5. Os fundos de mapa (CARTO/OSM/Esri) exigem internet.
