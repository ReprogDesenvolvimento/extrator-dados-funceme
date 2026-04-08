# 🌊 Extrator Anual de Tábua de Marés (Web Scraping)

![Excel](https://img.shields.io/badge/Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)
![Power Query](https://img.shields.io/badge/Power_Query-F2C811?style=for-the-badge&logo=power-bi&logoColor=black)
![Data Analysis](https://img.shields.io/badge/Análise_de_Dados-4285F4?style=for-the-badge&logo=google-analytics&logoColor=white)

> **Nota de Portfólio:** Este repositório atua como uma vitrine (Showcase). Os arquivos originais (`.xlsx`) não estão disponíveis para download público a fim de proteger a propriedade da ferramenta, mas toda a arquitetura, lógica e trechos de código estão documentados abaixo.

## 📋 Resumo do Projeto
Ferramenta automatizada construída 100% no Excel (via Power Query/Linguagem M) para extrair, limpar e consolidar o histórico anual completo de marés, fases da lua e nascer/pôr do sol do portal da FUNCEME, otimizando o tempo de coleta de dados para pesquisas acadêmicas e biológicas.

---

## ⚠️ O Problema
O portal da FUNCEME possui dados meteorológicos riquíssimos, mas a visualização na web é altamente fragmentada. A interface só exibe as informações em blocos de poucos dias por vez, misturando datas, horas e textos na mesma célula HTML. 

Para um pesquisador que precisa da base de um **ano inteiro**, fazer a extração e limpeza manualmente (semana a semana) seria um processo inviável, repetitivo e sujeito a erros.

<img width="1354" height="575" alt="site_extração" src="https://github.com/user-attachments/assets/16bf90bf-6055-4389-938d-933b15f1574f" />


## ⚙️ A Solução (Arquitetura Técnica)
Para entregar uma solução amigável ao usuário final (onde ele não precisasse instalar o Python ou rodar scripts complexos no terminal), decidi construir todo o pipeline de extração e tratamento diretamente no **Excel, utilizando Linguagem M**.

O fluxo de processamento (ETL) segue os seguintes passos:

1. **Estratégia de Varredura ("Pulo do Gato"):** O usuário digita o ano na planilha. O código M pega esse ano e cria uma lista de datas dinâmicas pulando de 5 em 5 dias, gerando as URLs exatas da requisição para cobrir os 365 dias com sobreposição.
2. **Web Scraping e Limpeza Pesada:** O script acessa as páginas, expande as tabelas HTML, separa dias da semana, utiliza a função `FillDown` para preencher células vazias estruturais e limpa strings de texto para isolar os horários.
3. **Enriquecimento de Dados:** Regras de negócio embutidas na consulta classificam automaticamente a maré como "Sizígia" (Nova/Cheia) ou "Quadratura" (Crescente/Minguante).
4. **Deduplicação Inteligente:** Filtro final de remoção de linhas sobrepostas utilizando a `Data` e a `Hora da Maré` como chaves primárias compostas.

### 💻 Trecho do Código (Linguagem M)
Abaixo, um trecho do motor principal responsável por gerar a lista de datas e fazer as requisições web dinâmicas:

```powerquery
let
    // 1. Puxa o ano definido pelo usuário
    GetAno = Excel.CurrentWorkbook(){[Name="TabelaAno"]}[Content]{0}[Ano],
    
    // 2. Estratégia de Varredura (O Pulo do Gato)
    DataInicial = Date.AddDays(#date(GetAno, 1, 1), -5),
    ListaDatas = List.Dates(DataInicial, 80, #duration(5, 0, 0, 0)), 
    TabelaDatas = Table.FromList(ListaDatas, Splitter.SplitByNothing(), {"DataBusca"}),
    
    // 3. Busca no Site (Requisição dinâmica por URL)
    AdicionarURL = Table.AddColumn(TabelaDatas, "URL", each "[https://tabuadasmares.funceme.br/2/](https://tabuadasmares.funceme.br/2/)" & Date.ToText([DataBusca], "yyyy-MM-dd") & "/true/-3.54826/-38.81478/true/true"),
    PuxarDados = Table.AddColumn(AdicionarURL, "Conteudo", each try Web.Page(Web.Contents([URL])) otherwise null)
    
    // [O fluxo de código continua com Expansão HTML, FillDown, Limpeza de Strings e Tipagem...]
in
    PuxarDados
```

---

## ✅ O Resultado

O produto final é uma interface limpa de "1 clique". O usuário altera a célula do ano, clica em "Atualizar Tudo" e o algoritmo consome as páginas em segundo plano, entregando um banco de dados completo, tipado, ordenado e pronto para ser consumido em análises ou dashboards.

<img width="1360" height="660" alt="Excel_mares" src="https://github.com/user-attachments/assets/4e80739a-51e2-4ac4-b8b2-d4b346522b3f" />

---
**Desenvolvido por Amanda Lira**  
*Automação de processos usando as ferramentas certas para democratizar o acesso à informação.*
