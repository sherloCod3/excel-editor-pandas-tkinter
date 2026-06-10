# Relatorio tecnico academico: SecOps Portal

## 1. Resumo executivo

O SecOps Portal e uma aplicacao web em Python e Streamlit para apoio a operacoes de seguranca. O sistema permite carregar logs em CSV ou Excel, editar os dados na interface, exportar o resultado e investigar IPs ou dominios por meio de um fluxo automatizado no n8n. A investigacao consulta fontes externas, como VirusTotal e ip-api.com, e pode gerar um resumo em linguagem natural com um LLM compativel com a API da OpenAI ou com Ollama local.

O projeto nasceu como uma aplicacao desktop em Tkinter para manipulacao de planilhas Excel com Pandas. Durante o desenvolvimento, o escopo foi reconstruido para um portal SOC, porque a manipulacao de planilhas isolada nao representava bem um fluxo real de analise de seguranca. Escolhemos Streamlit porque ele permitiu manter a base em Python e entregar uma interface web reativa com menos codigo de infraestrutura. Descartamos uma arquitetura Flask ou FastAPI com React porque ela exigiria mais camadas para resolver uma experiencia que, neste projeto, precisava ser orientada a dados e prototipagem controlada.

A principal contribuicao tecnica esta na separacao entre interface, logica de negocio e automacao externa. O modulo `core.py` valida respostas, aplica defaults e limita o score de ameaca antes da renderizacao. O workflow n8n isola a orquestracao SOC, incluindo branching para IP ou dominio, consultas HTTP, enriquecimento opcional com IA e resposta em JSON. A suite de testes cobre parsing de arquivos, modo mock, modo live simulado, validacao de schema e integracao com n8n quando o webhook esta configurado.

## 2. Introducao

Em ambientes pequenos, laboratorios academicos ou equipes sem SIEM dedicado, a investigacao de logs costuma comecar em planilhas. Esse fluxo e acessivel, mas cria problemas. O analista precisa alternar entre arquivo local, navegador, consultas manuais em bases externas e anotacoes soltas. Essa alternancia aumenta o tempo de investigacao e facilita erros de interpretacao, principalmente quando o dado de entrada vem de diferentes formatos.

O projeto iniciou como uma ferramenta desktop em Tkinter para editar planilhas Excel usando Pandas. Essa origem explica a primeira necessidade tecnica: carregar dados tabulares, permitir ajustes e exportar o resultado. Durante a evolucao do escopo, identificamos que a manipulacao de dados ficaria mais relevante se fosse conectada a um caso de uso de seguranca. Por isso, a aplicacao foi reconstruida como um portal SOC web, com interface Streamlit, tema de terminal e fluxo de investigacao de IPs e dominios.

O objetivo geral e construir um portal de apoio a analise operacional de seguranca que funcione em dois contextos: demonstracao offline e investigacao conectada a um backend de automacao. O modo offline usa dados simulados para manter a aplicacao utilizavel sem chaves ou infraestrutura externa. O modo conectado envia o alvo para um webhook n8n, coleta inteligencia de ameacas e devolve um schema padronizado para a interface.

O projeto tambem tem um objetivo pedagogico. Ele mostra como uma aplicacao simples de manipulacao de planilhas pode evoluir para uma arquitetura com validacao, automacao, testes e gestao de credenciais. Essa evolucao nao e apenas visual. Ela muda a responsabilidade do sistema: de editor de dados para uma ferramenta que recebe entrada de usuario, chama servicos externos e precisa tratar falhas, latencia e limites de seguranca.

## 3. Objetivos especificos

O primeiro objetivo especifico foi manter o fluxo de analise de logs acessivel. Para isso, o sistema aceita arquivos CSV, XLS e XLSX, transforma o conteudo em um `DataFrame` do Pandas e apresenta os dados em uma tabela editavel. Escolhemos Pandas porque ele e adequado para dados tabulares e ja oferece funcoes de leitura para CSV e Excel. Descartamos implementar parsers proprios porque isso aumentaria risco de erro em formatos ja bem suportados pela biblioteca.

O segundo objetivo foi adicionar investigacao de ameacas para IPs e dominios. O usuario informa um alvo, e a aplicacao decide entre MOCK MODE e LIVE MODE conforme a presenca de URL de webhook. Escolhemos esse chaveamento porque ele permite testar e demonstrar a interface sem depender de VirusTotal, n8n ou chave de LLM. Descartamos exigir o backend desde a primeira execucao porque isso dificultaria a avaliacao academica e o uso em sala.

O terceiro objetivo foi padronizar a resposta antes da renderizacao. A funcao `_validate_response()` garante campos esperados, converte `threat_score` para inteiro e limita o valor entre 0 e 100. Escolhemos validar no backend Python da interface porque o n8n pode sofrer alteracoes de workflow e retornar tipos inesperados. Descartamos confiar diretamente no JSON recebido porque uma resposta incompleta poderia quebrar o medidor visual ou gerar conclusoes incorretas.

O quarto objetivo foi criar uma suite de testes que protegesse o comportamento principal. Os testes unitarios cobrem parsing, modo mock, live mode com `requests.post` mockado, erros HTTP e regressao do timeout de 15 segundos. Os testes de integracao validam o contrato com n8n quando `N8N_WEBHOOK_URL` existe. Escolhemos auto-skip para integracao porque o teste depende de infraestrutura externa. Descartamos mocks totais como unica estrategia porque eles nao provam que Streamlit e n8n concordam sobre o schema real.

## 4. Requisitos

### 4.1 Requisitos funcionais

O sistema deve permitir upload de arquivos CSV, XLS e XLSX contendo logs de rede ou listas de IPs. Depois do upload, a aplicacao deve carregar o conteudo em uma tabela editavel e permitir exportacao em CSV. Esse requisito atende ao caso de uso original do projeto: manipular dados tabulares sem exigir que o usuario escreva codigo Python.

O sistema deve permitir investigacao de um IP, dominio ou URL na tela THREAT_INTEL. Quando nao houver webhook configurado, ele deve retornar uma resposta simulada com score, localizacao e status. Quando houver webhook, ele deve enviar um POST para o n8n com `{"target": "<valor>"}` e receber uma resposta em JSON. Escolhemos esse contrato minimo porque reduz acoplamento entre a interface e o workflow.

O sistema deve apresentar o resultado com um medidor de ameaca, cards de geolocalizacao, status malicioso ou limpo, resposta bruta em JSON e historico recente. O resumo de IA deve aparecer apenas quando existir conteudo retornado pelo workflow. Escolhemos condicionar a exibicao do card de IA porque o LLM e opcional. Descartamos mostrar um espaco vazio permanente porque isso confundiria o usuario em execucoes sem chave.

O sistema deve salvar a URL do webhook em `config.json`, com permissao de leitura e escrita apenas para o dono do arquivo. Esse requisito facilita o uso local sem pedir a URL a cada execucao. Ao mesmo tempo, ele reconhece que a URL pode conter token ou caminho sensivel. Em deploy no Streamlit Cloud, a estrategia recomendada muda para Streamlit Secrets, porque o filesystem da plataforma nao deve ser tratado como armazenamento permanente de configuracao sensivel.

### 4.2 Requisitos nao funcionais

O sistema deve responder de forma previsivel mesmo quando servicos externos falham. A funcao `fetch_soc_data()` captura excecoes de `requests`, retorna uma chave `error` e evita que a interface quebre. O timeout de 15 segundos foi protegido por teste porque VirusTotal e chamadas encadeadas no n8n podem ser mais lentos que uma chamada HTTP simples. Escolhemos timeout explicito porque chamadas sem limite podem prender a experiencia do usuario.

O sistema deve proteger credenciais. Chaves de VirusTotal e OpenAI ficam em variaveis de ambiente no n8n e nao sao versionadas. A URL do webhook local e salva em arquivo gitignored, com permissao restrita. Escolhemos variaveis de ambiente porque elas separam segredo de codigo. Descartamos hardcode porque qualquer commit ou compartilhamento do repositorio poderia expor credenciais.

O sistema deve ser testavel sem infraestrutura externa. O MOCK MODE viabiliza testes e demonstracoes sem n8n. Os testes de integracao usam auto-skip quando `N8N_WEBHOOK_URL` nao esta definido. Escolhemos essa separacao porque o feedback rapido dos testes unitarios e mais importante no ciclo diario, enquanto integracao deve ser executada quando o backend estiver disponivel.

O sistema deve ser portavel entre ambiente local, Streamlit Community Cloud e uma VPS para n8n. A interface roda como aplicacao Python com dependencias em `requirements.txt`. O n8n roda via Docker Compose. Escolhemos deploy dividido porque Streamlit Cloud resolve bem a interface, mas nao hospeda o backend de automacao com persistencia e credenciais do mesmo modo. Descartamos colocar tudo em um unico processo porque isso misturaria responsabilidades e complicaria operacao.

## 5. Arquitetura e decisoes tecnicas

A arquitetura divide o sistema em tres partes: interface Streamlit, logica Python em `core.py` e workflow n8n. A interface cuida de navegacao, upload, entrada do alvo e visualizacao. O `core.py` concentra leitura de arquivos, chamada HTTP e validacao da resposta. O n8n executa a orquestracao SOC. Escolhemos essa divisao porque cada parte tem uma responsabilidade clara. Descartamos colocar chamadas de VirusTotal diretamente na interface porque isso aumentaria acoplamento e dificultaria evoluir o fluxo.

Escolhemos Streamlit porque a aplicacao e orientada a dados e escrita por um desenvolvedor Python. O framework oferece componentes como upload, editor de dados, botao de download e estado de sessao sem exigir frontend separado. Descartamos Flask para este projeto porque ele entregaria rotas HTTP, mas exigiria criar templates ou frontend manual para a experiencia reativa. Descartamos FastAPI com React porque a combinacao seria adequada para produto maior, mas traria build frontend, API propria, CORS e mais pontos de teste para um escopo academico.

Escolhemos n8n como orquestrador SOC porque ele permite visualizar e alterar o fluxo de chamadas externas sem recompilar a aplicacao. O workflow tambem torna claro o branching entre IP e dominio, o enriquecimento por geolocalizacao e VirusTotal, e a decisao de usar ou nao IA. Descartamos scripts Python diretos para toda a orquestracao porque eles deixariam o fluxo menos auditavel para demonstracao e concentrariam credenciais, branching e chamadas externas no mesmo processo da interface.

Escolhemos MOCK MODE e LIVE MODE como estrategia de testabilidade e demonstracao. O MOCK MODE devolve dados coerentes com o schema esperado quando nao existe webhook configurado. O LIVE MODE e ativado automaticamente quando o usuario salva a URL do webhook. Descartamos uma tela de configuracao obrigatoria porque ela impediria o uso inicial e criaria dependencia de chaves externas antes mesmo de avaliar a interface.

Escolhemos permitir LLM compativel com OpenAI e citar Ollama como alternativa local. A decisao reduz custo e melhora privacidade quando o resumo de ameaca nao deve sair da maquina do usuario. Ollama oferece compatibilidade experimental com endpoints da API OpenAI, o que permite reaproveitar clientes e nos de integracao que esperam esse formato. Descartamos depender apenas da OpenAI porque isso tornaria a funcionalidade de resumo condicionada a conta, saldo e envio de dados a um provedor externo.

Escolhemos pytest com auto-skip para testes de integracao. A documentacao do pytest permite pular testes quando uma condicao externa nao esta satisfeita, e o projeto aplica isso para `N8N_WEBHOOK_URL`. Descartamos transformar todos os testes em mocks porque isso nao validaria o contrato real entre Streamlit e n8n. Tambem descartamos exigir n8n para toda execucao porque isso tornaria a suite lenta e dependente de Docker em qualquer ambiente.

Escolhemos Docker para executar n8n porque o container encapsula runtime, dependencias e variaveis de ambiente. O Docker Compose permite declarar servico, porta e ambiente em um arquivo versionavel. Descartamos instalacao manual do n8n porque ela depende de versao local de Node.js, permissao de sistema e configuracao repetida em cada maquina. Em uma VPS, Compose tambem facilita restart e reproducao do ambiente.

O workflow n8n importavel possui 15 nos no arquivo `n8n/soc_agent_workflow.json`: webhook, sanitizacao, decisao IP ou dominio, chamadas HTTP, merges, construcao de schema, feature flag de IA, modelo OpenAI, mensagem de IA desabilitada e resposta final. Em descricao conceitual, alguns desses nos podem ser agrupados em 12 a 14 etapas porque o no `OpenAI Chat Model` funciona como suporte ao `AI - Generate Threat Summary`. No relatorio, adotamos a contagem tecnica do arquivo importavel e explicamos o agrupamento para evitar divergencia entre documentacao e artefato.

## 6. Seguranca

A primeira decisao de seguranca foi nao colocar credenciais no codigo. A chave do VirusTotal e a chave OpenAI ficam em variaveis de ambiente no n8n. No Streamlit Cloud, a recomendacao e usar Streamlit Secrets, pois a propria documentacao do Streamlit orienta armazenar segredos fora do repositorio e permite configura-los na plataforma. Escolhemos essa separacao porque codigo-fonte e segredo tem ciclos de vida diferentes. Descartamos `.env` versionado porque ele transforma uma configuracao local em vazamento persistente.

A URL do webhook e tratada como sensivel. Localmente, ela e salva em `config.json`, arquivo ignorado pelo Git, e o codigo aplica `chmod` para restringir leitura e escrita ao usuario dono. Essa medida nao substitui autenticacao no n8n, mas reduz exposicao acidental em maquina compartilhada. Escolhemos persistir apenas essa URL localmente porque ela melhora usabilidade. Descartamos salvar chaves de API na interface porque o frontend Streamlit nao precisa conhecer credenciais dos fornecedores externos.

O `core.py` contem uma validacao de URL de webhook contra SSRF basico. SSRF, ou Server-Side Request Forgery, ocorre quando uma entrada controlada pelo usuario faz o servidor consultar enderecos internos ou metadados de infraestrutura. O codigo bloqueia esquemas que nao sejam HTTP ou HTTPS e hosts como `localhost`, `127.*`, `0.0.0.0`, `169.254.*` e loopback IPv6. Escolhemos essa barreira porque o webhook e informado pelo usuario. Descartamos aceitar qualquer URL porque isso poderia transformar a aplicacao em proxy para rede interna.

No n8n, o no Set extrai e sanitiza o campo `target` antes das chamadas externas. Essa etapa limita a superficie de entrada do workflow e reduz a chance de o valor informado ser usado sem normalizacao. Escolhemos sanitizar cedo porque o mesmo alvo alimenta ramificacoes e requisicoes HTTP. Descartamos sanitizar apenas no final porque o dano ocorreria antes da montagem da resposta.

O projeto envia dados a servicos externos no LIVE MODE. VirusTotal recebe IPs ou URLs para analise; ip-api.com recebe IPs para geolocalizacao; o LLM recebe o contexto usado para gerar o resumo. Isso deve ser explicado ao usuario, principalmente quando os logs contiverem indicadores internos ou sensiveis. Escolhemos manter a IA opcional porque a funcao de resumo agrega legibilidade, mas tambem amplia a exposicao de dados. Descartamos ativar IA por padrao porque isso criaria compartilhamento externo sem decisao explicita.

As limitacoes conhecidas incluem ausencia de autenticacao propria na aplicacao Streamlit, dependencia de politicas do n8n para proteger o webhook em producao, limite de requisicoes do VirusTotal free tier e confianca parcial no resultado de terceiros. O score de ameaca e um apoio visual, nao uma prova definitiva. Falsos positivos e falsos negativos podem ocorrer em feeds de reputacao. Por isso, o sistema mostra a resposta bruta e detalhes, permitindo revisao manual.

## 7. Implementacao

O modulo LOG_ANALYSIS fica na funcao `show_data_editor()` de `app.py`. Ele usa `st.file_uploader` para receber arquivos, `load_data()` para converter o conteudo em `DataFrame`, `st.data_editor` para edicao inline e `st.download_button` para exportacao CSV. Escolhemos usar componentes nativos do Streamlit porque eles reduzem codigo de interface e mantem o fluxo no Python. Descartamos construir uma grade JavaScript separada porque isso exigiria ponte entre frontend e backend sem necessidade para o escopo.

A funcao `load_data()` aceita CSV, XLS e XLSX. Ela escolhe `pd.read_csv` para CSV e `pd.read_excel` para planilhas. Se o formato nao for suportado ou ocorrer erro de leitura, retorna `None`. Escolhemos esse retorno simples porque a interface consegue exibir erro sem expor stack trace ao usuario. Descartamos propagar excecoes para a tela porque arquivos de log podem ser enviados com delimitadores, extensoes ou conteudo incorretos.

```python
def load_data(file) -> pd.DataFrame | None:
    """Parse an uploaded CSV or Excel file into a DataFrame."""
    try:
        if file.name.endswith(".csv"):
            return pd.read_csv(file)
        elif file.name.endswith((".xls", ".xlsx")):
            return pd.read_excel(file)
        return None
    except Exception:
        return None
```

O modulo THREAT_INTEL fica em `show_soc_investigator()`. A tela carrega a configuracao, recebe o alvo, executa uma animacao curta de scan e chama `fetch_soc_data()`. Se a resposta contem `error`, mostra erro de conexao. Se a resposta e valida, atualiza contador, historico, medidor de ameaca, card de IA, geolocalizacao, status e JSON bruto. Escolhemos manter historico limitado a cinco itens porque isso ajuda comparacao sem crescer indefinidamente em `st.session_state`.

O fluxo completo do scan com LIVE MODE e: usuario informa alvo, Streamlit envia POST ao webhook, n8n sanitiza `target`, decide se e IP ou dominio, consulta ip-api.com e VirusTotal conforme o caso, normaliza o schema, opcionalmente chama LLM, monta JSON final e devolve ao Streamlit. A interface entao chama `_validate_response()` antes de renderizar. Escolhemos essa validacao mesmo depois do n8n porque workflows visuais podem ser alterados durante manutencao. Descartamos confiar apenas no contrato documentado porque contrato sem guard rail em runtime nao protege a UI.

```python
def _validate_response(data: dict[str, Any]) -> dict[str, Any]:
    defaults: dict[str, Any] = {
        "status": "success",
        "target": "",
        "threat_score": 0,
        "location": "Unknown",
        "known_malicious": False,
        "summary": "",
        "details": "",
    }
    merged = {**defaults, **data}
    try:
        merged["threat_score"] = max(0, min(100, int(merged["threat_score"])))
    except (ValueError, TypeError):
        merged["threat_score"] = 0
    return merged
```

A funcao `fetch_soc_data()` implementa o chaveamento entre mock e live. Sem webhook, ela aguarda 1,5 segundo para simular processamento e devolve uma resposta aleatoria dentro do schema. Com webhook, valida a URL, executa `requests.post` com timeout de 15 segundos e valida o JSON recebido. Escolhemos esse desenho porque a mesma interface consome mock e live sem codigo duplicado. Descartamos criar duas telas separadas porque isso multiplicaria estados e testes.

O workflow n8n possui 15 nos no arquivo importavel. A descricao operacional e: 1) recebe requisicao; 2) sanitiza alvo; 3) decide IP ou dominio; 4) consulta geolocalizacao para IP; 5) consulta VirusTotal para IP; 6) consulta VirusTotal para dominio ou URL; 7) combina resultados; 8) normaliza schema; 9) verifica se a chave de IA existe; 10) gera resumo; 11) fornece modelo OpenAI ao no de IA; 12) cria mensagem quando IA esta desabilitada; 13) anexa resumo; 14) monta resposta final; 15) responde ao webhook. Escolhemos branching explicito porque a consulta de IP e dominio usa endpoints e dados diferentes. Descartamos um unico bloco generico porque isso esconderia diferencas de tratamento e dificultaria depuracao.

## 8. Testes

A estrategia de testes separa unidade e integracao. Os testes unitarios rodam sem n8n, sem VirusTotal e sem OpenAI. Eles validam funcoes Python com entradas controladas e mocks de rede. Os testes de integracao dependem de `N8N_WEBHOOK_URL` e validam o contrato real com o workflow. Escolhemos essa divisao porque falhas de codigo local e falhas de infraestrutura precisam ser diagnosticadas de formas diferentes.

O arquivo `test_validate_response.py` protege o schema recebido pela interface. Ele cobre passagem de resposta valida, conversao de string para inteiro, conversao de float, tratamento de `None`, valores invalidos, clamping acima de 100, clamping abaixo de 0, defaults para campos ausentes e preservacao de campos extras. Essa cobertura e importante porque a tela de resultado depende de `threat_score`, `location` e `known_malicious` para decidir cores, textos e cards.

O arquivo `test_load_data.py` protege leitura de CSV e Excel. Ele usa objetos em memoria que simulam arquivos enviados pelo Streamlit. Os testes verificam CSV valido, XLSX valido, extensao nao suportada e tratamento de erros. Escolhemos testes em memoria porque eles reduzem dependencia de arquivos temporarios e deixam o comportamento mais direto. Descartamos testar somente pela interface porque o problema principal esta no parsing.

O arquivo `test_fetch_soc_data.py` protege o comportamento de MOCK MODE e LIVE MODE simulado. Ele verifica que o mock retorna dicionario, campos obrigatorios, alvo ecoado, score inteiro dentro de faixa e ausencia de erro. No live simulado, verifica chamada ao webhook, payload, passagem de summary, coercao de score string e erros de conexao, timeout e HTTP. O teste de timeout garante 15 segundos, evitando regressao para valores menores que falhavam em latencias maiores.

O arquivo `test_integration_n8n.py` valida a integracao quando o backend esta disponivel. Ele deve ser executado com o n8n ativo e `N8N_WEBHOOK_URL` definido. Quando a variavel nao existe, `conftest.py` aplica skip automatico aos testes marcados como integracao. Isso nao e uma limitacao, mas uma decisao arquitetural. O sistema reconhece que dependencia externa nao deve impedir a execucao da suite local.

O auto-skip melhora a confiabilidade do ciclo de desenvolvimento. Sem ele, qualquer ambiente sem Docker, webhook ou chaves falharia antes de testar o codigo Python. Com ele, a suite unitária continua verificando 36 testes locais, enquanto a integracao e acionada quando a infraestrutura existe. Escolhemos esse comportamento porque ele separa disponibilidade de servico externo de qualidade do codigo local.

## 9. Deploy

A estrategia de deploy e dividida. O frontend Streamlit pode ser publicado no Streamlit Community Cloud, puxando o codigo direto do GitHub. O backend n8n deve rodar em uma VPS Linux com Docker, como DigitalOcean, Hetzner, AWS EC2, Railway ou Render, conforme restricoes de custo e persistencia. Escolhemos essa separacao porque a interface e o orquestrador tem necessidades diferentes de runtime.

O Streamlit Cloud e adequado para a camada visual porque a aplicacao e Python, tem dependencias declaradas e nao exige servidor customizado. No entanto, o filesystem da plataforma nao deve ser usado como armazenamento permanente de segredos. Por isso, credenciais devem ir para Streamlit Secrets quando forem usadas pela aplicacao. No desenho atual, as chaves principais ficam no n8n, e a URL do webhook deve ser tratada como configuracao sensivel.

O n8n em Docker concentra automacao, credenciais e chamadas externas. Em producao, ele deve estar atras de HTTPS, com autenticacao, variaveis de ambiente protegidas e backups do volume de dados. Escolhemos Docker Compose porque o arquivo `docker-compose.yml` declara a execucao de forma reproduzivel. Descartamos instalacao manual porque ela dificulta repetir o ambiente em outra maquina e aumenta divergencia entre desenvolvimento e producao.

O MOCK MODE e parte da estrategia de portabilidade. Mesmo sem n8n, VirusTotal ou OpenAI, o aplicativo abre no Streamlit Cloud e permite demonstrar navegacao, upload, edicao e visualizacao de um scan simulado. Isso e importante em apresentacoes academicas, onde a rede pode falhar ou chaves de API podem nao estar disponiveis. Escolhemos manter essa funcionalidade em producao porque ela tambem serve como modo de degradacao controlada.

## 10. Desafios e decisoes durante o desenvolvimento

O primeiro desafio foi transformar uma aplicacao desktop de planilhas em um portal web de seguranca sem perder a funcionalidade original. A solucao foi manter a manipulacao tabular como modulo LOG_ANALYSIS e adicionar THREAT_INTEL como novo modulo. Escolhemos evoluir o escopo em camadas porque isso preservou valor do projeto inicial. Descartamos abandonar o editor de logs porque ele ainda representa a entrada natural de muitos fluxos de investigacao.

O segundo desafio foi lidar com respostas externas instaveis. APIs podem retornar campos ausentes, tipos diferentes, erros HTTP ou respostas lentas. A solucao foi centralizar `_validate_response()` e testar coercao, defaults e clamping. Escolhemos validar antes de renderizar porque a interface visual depende de tipos previsiveis. Descartamos tratar cada campo individualmente na UI porque isso espalharia regras e aumentaria risco de inconsistencia.

O terceiro desafio foi permitir demonstracao sem infraestrutura. O projeto depende de n8n, VirusTotal e opcionalmente LLM, mas exigir todos esses componentes reduziria a capacidade de avaliacao. A solucao foi criar MOCK MODE automatico quando nao ha webhook. Escolhemos ativacao por ausencia de URL porque e simples para o usuario. Descartamos um arquivo de configuracao obrigatorio porque ele criaria friccao no primeiro uso.

O quarto desafio foi proteger o webhook sem bloquear o uso local. O sistema salva a URL em `config.json`, mas aplica permissao restrita. Alem disso, `core.py` bloqueia destinos internos para reduzir risco de SSRF quando uma URL e fornecida pelo usuario. Escolhemos controles simples e verificaveis porque o escopo e academico, mas envolve chamadas reais de rede. Descartamos ignorar esse risco porque qualquer aplicacao que faz POST para URL configuravel precisa tratar abuso.

O quinto desafio foi conciliar documentacao e artefato real do workflow. O briefing menciona 12 nos, o README do n8n descreve 14 etapas e o JSON importavel possui 15 nos. A diferenca ocorre porque algumas descricoes agrupam o modelo OpenAI com a cadeia de IA, enquanto o n8n registra ambos como nos separados. Escolhemos documentar a contagem tecnica de 15 nos e explicar o agrupamento conceitual. Descartamos ajustar o texto para parecer uniforme porque isso ocultaria uma divergencia rastreavel.

## 11. Resultados e validacao

O resultado e uma aplicacao funcional com duas areas principais. LOG_ANALYSIS permite carregar dados, editar registros e exportar CSV. THREAT_INTEL permite informar alvo, executar uma investigacao simulada ou real, visualizar score de ameaca, geolocalizacao, status, resumo de IA e resposta bruta. A interface usa estilo cyberpunk e terminal, mas a estrutura de dados permanece simples e rastreavel.

A validacao automatizada cobre os principais riscos do codigo Python. Segundo a organizacao do projeto, ha 36 testes unitarios distribuidos entre validacao de resposta, carregamento de dados e fetch de dados SOC. Os testes de integracao ficam disponiveis para verificar o contrato n8n quando o webhook existe. Esse conjunto valida que o sistema nao depende apenas de teste manual na interface.

A validacao arquitetural aparece no proprio comportamento do modo mock e live. Quando nao ha webhook, a aplicacao continua operando e demonstra o fluxo visual. Quando ha webhook, ela muda para LIVE MODE e executa a cadeia externa. Esse comportamento confirma que a interface nao esta rigidamente presa a um backend. Ao mesmo tempo, `_validate_response()` confirma que o backend nao tem permissao implicita para quebrar a renderizacao com tipos fora do esperado.

Os resultados tambem mostram limitacoes. O score depende da qualidade dos dados de VirusTotal e da logica de normalizacao no workflow. A geolocalizacao de IP nao identifica necessariamente o atacante real, apenas metadados do endereco consultado. O resumo de IA deve ser lido como apoio interpretativo, nao como decisao automatica. Essas limitacoes foram mantidas visiveis para evitar que a ferramenta seja apresentada como substituta de analise humana.

## 12. Conclusao e proximos passos

O SecOps Portal atingiu o objetivo de evoluir uma ferramenta de planilhas para uma aplicacao web de apoio a operacoes de seguranca. A solucao combina manipulacao de logs, investigacao de indicadores, automacao com n8n, resumo opcional com LLM e testes automatizados. As principais decisoes tecnicas foram guiadas por simplicidade operacional, separacao de responsabilidades e capacidade de demonstracao sem infraestrutura externa.

As escolhas feitas tem justificativas claras. Escolhemos Streamlit porque o projeto e orientado a dados e Python. Escolhemos n8n porque o fluxo SOC fica visivel e alteravel. Escolhemos Docker porque o backend precisa de ambiente reproduzivel. Escolhemos pytest com auto-skip porque testes locais e integracao externa nao devem ter a mesma exigencia de ambiente. Escolhemos Ollama como alternativa porque privacidade e custo importam quando o resumo usa dados de seguranca.

Como proximos passos, o projeto pode adicionar autenticacao de usuario na interface, controle de papeis no n8n, armazenamento historico de scans, exportacao de relatorios em PDF e suporte a mais fontes de inteligencia de ameacas. Tambem seria util adicionar validacao formal de entrada para distinguir IP, dominio e URL antes do envio ao webhook. Outra melhoria seria registrar metricas de latencia e falhas para avaliar o comportamento do fluxo em uso continuo.

Tambem e recomendavel endurecer a implantacao em producao. O n8n deve rodar atras de proxy HTTPS, com credenciais fortes, backups, atualizacoes planejadas e restricao de rede para o webhook. A interface deve evitar expor detalhes sensiveis em mensagens de erro. O pipeline de testes pode incluir uma etapa de integracao em CI quando um ambiente n8n de teste estiver disponivel. Com essas evolucoes, o projeto ficaria mais preparado para uso fora do contexto academico.

## 13. Referencias

Documentacao oficial do Streamlit. "Secrets management". Disponivel em: https://docs.streamlit.io/develop/concepts/connections/secrets-management. Acesso em: 10 jun. 2026.

Documentacao oficial do Streamlit. "`st.data_editor`". Disponivel em: https://docs.streamlit.io/develop/api-reference/data/st.data_editor. Acesso em: 10 jun. 2026.

Documentacao oficial do Docker. "Docker Compose". Disponivel em: https://docs.docker.com/compose/. Acesso em: 10 jun. 2026.

Documentacao oficial do n8n. "Docker installation". Disponivel em: https://docs.n8n.io/hosting/installation/docker/. Acesso em: 10 jun. 2026.

Documentacao oficial do VirusTotal. "VirusTotal API v3 Overview". Disponivel em: https://docs.virustotal.com/reference/overview. Acesso em: 10 jun. 2026.

Documentacao oficial do pytest. "How to use skip and xfail to deal with tests that cannot succeed". Disponivel em: https://docs.pytest.org/en/stable/how-to/skipping.html. Acesso em: 10 jun. 2026.

Documentacao oficial do Pandas. "`pandas.read_csv`". Disponivel em: https://pandas.pydata.org/docs/reference/api/pandas.read_csv.html. Acesso em: 10 jun. 2026.

Documentacao do Ollama. "OpenAI compatibility". Disponivel em: https://ollama.readthedocs.io/en/openai/. Acesso em: 10 jun. 2026.

Repositorio do projeto SecOps Portal. Arquivos consultados: `README.md`, `app.py`, `core.py`, `n8n/README.md`, `n8n/soc_agent_workflow.json`, `tests/conftest.py`, `tests/test_validate_response.py`, `tests/test_fetch_soc_data.py` e `pyproject.toml`. Acesso local em: 10 jun. 2026.
