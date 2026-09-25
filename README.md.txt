Gestão de Hora Extra 2.1

Automatização do processo de gestão de hora extra utilizando Microsoft Forms, Microsoft Lists, SharePoint, Power Query e VBA.

📌 Visão Geral

Gestão de Hora Extra 2.1 é uma solução desenvolvida para automatizar e centralizar o processo de gestão de hora extra utilizando recursos do ecossistema Microsoft 365.

O projeto surgiu como uma evolução direta do projeto anterior hour-extra-automation, substituindo processos baseados exclusivamente em Excel por uma arquitetura integrada com Microsoft Forms, Microsoft Lists, SharePoint, Power Query e VBA.

O objetivo é reduzir atividades manuais, aumentar a confiabilidade das informações, melhorar a experiência dos usuários e fornecer maior controle operacional para analistas e supervisores.

🖼️ Demonstração da Solução
Fluxo Geral da Arquitetura

IMAGES/01-Sharepoint.png

A solução integra Microsoft Forms, Microsoft Lists, SharePoint, Power Query e Excel em um fluxo único de gestão operacional.

Controle de Disponibilidade

IMAGES/10-Disponibilidade.png

Controle centralizado da disponibilidade dos colaboradores para realização de Hora Extra.

Gestão de Solicitações

IMAGES/13-Solicitações.png

Registro e gerenciamento das solicitações de Hora Extra realizadas pelos analistas.

Área Operacional de Hora Extra

IMAGES/17-HoraExtra-suja.png

Área operacional utilizada para consolidação, tratamento e distribuição das atividades.

Identificação de Conflitos

IMAGES/24-ConflitosAtualizados.png

Detecção automática de conflitos de alocação entre colaboradores e atividades.

Painel Consolidado

IMAGES/15-ResumoHoras.png

Painel de indicadores consolidando horas, origens, centros de custo e recursos utilizados.

🎯 Problema de Negócio

O processo anterior dependia de múltiplas planilhas e diversas etapas manuais.

Entre os principais desafios estavam:

Digitação manual de IPNs;
Possibilidade de erros na seleção de colaboradores;
Necessidade de consultas externas para localizar funcionários;
Controle descentralizado da disponibilidade;
Dificuldade para identificar conflitos entre atividades;
Maior risco de inconsistências nos dados.

Mesmo com automações existentes, a qualidade das informações ainda dependia significativamente das ações dos usuários.

🚀 Solução Desenvolvida

A solução foi reformulada para utilizar os recursos do Microsoft 365 como principal fonte de informações.

Os formulários passaram a alimentar listas corporativas no Microsoft Lists, que são consumidas pelo Power Query e posteriormente processadas pelo Excel através de automações em VBA.

Com isso, o Excel passou a atuar como plataforma operacional e analítica, enquanto o armazenamento e gerenciamento das informações ficaram centralizados no SharePoint e Microsoft Lists.

📈 Evolução da Arquitetura
Versão 1.0

Microsoft Forms

↓

Planilhas de Resposta

↓

Power Query

↓

Solicitações

↓

VBA

↓

Hora Extra

Versão 2.1

Microsoft Forms

↓

Microsoft Lists

↓

SharePoint

↓

Power Query

↓

Excel Operacional

↓

VBA

↓

Indicadores e Relatórios

🔄 Fluxo Operacional
Disponibilidade

Colaborador

↓

Microsoft Forms

↓

Microsoft Lists

↓

Power Query

↓

Disponibilidade

↓

Validações e Conflitos

Solicitações

Analista

↓

Microsoft Forms

↓

Microsoft Lists

↓

Power Query

↓

Solicitações Hora Extra

↓

Importação Automática

↓

Hora Extra

↓

Resumo de Horas

🛠️ Tecnologias Utilizadas
Microsoft Forms
Microsoft Lists
SharePoint
Excel
Power Query
VBA
Microsoft 365
📋 Principais Funcionalidades
✅ Gestão de Disponibilidade

Permite registrar:

Data disponível;
Turno disponível;
Quantidade de horas disponíveis;
Necessidade de transporte;
Aprovação do supervisor.
✅ Gestão de Solicitações

Permite registrar e controlar:

Criticidade;
Impacto operacional;
Ganho esperado;
Origem da atividade;
Centro de custo;
Tempo de execução;
Colaboradores sugeridos.
✅ Cadastro de Colaboradores

Cadastro centralizado contendo:

IPN;
Nome;
E-mail;
Turno;
Função;
Supervisor;
BPU.
✅ Importação Automática

Importa atividades da área de solicitações para a área operacional de Hora Extra, reduzindo atividades manuais e garantindo maior rastreabilidade.

✅ Chave de Controle

Cada registro recebe uma chave única composta por:

ID_ATIVIDADE | IPN

Exemplos:

113 | IP00002

113 | IP00003

127 | IP00009

Essa chave representa a relação entre atividade e colaborador, permitindo rastreabilidade, controle de duplicidades e integridade durante as importações.

✅ Identificação de Conflitos

O sistema identifica automaticamente quando um colaborador está vinculado a mais de uma atividade na mesma data.

Os registros são destacados visualmente para facilitar a tomada de decisão pelos analistas e evitar alocação incorreta de recursos.

✅ Resumo de Horas

Painel consolidado contendo:

Horas por data;
Horas por origem;
Horas por centro de custo;
Horas por colaborador;
Horas por classificação MOD/MOS.
✅ Recuperação Operacional

Através da aba Formatação Original, o sistema consegue restaurar rapidamente a estrutura operacional da aba Hora Extra.

✅ Verificação do Sistema

Área destinada ao diagnóstico operacional da solução.

Permite identificar problemas relacionados a:

Eventos desabilitados;
Atualizações interrompidas;
Comportamentos inesperados;
Falhas operacionais.
✅ Barra de Progresso

Interface desenvolvida em VBA para acompanhamento visual da execução de atualizações e processos automáticos.

📈 Evolução da Versão 1.0 para 2.1
Antes
Forms + Excel;
Digitação manual de IPNs;
Dependência de múltiplas planilhas;
Maior risco de erro humano;
Controle operacional distribuído;
Ausência de identificação automática de conflitos.
Depois
Forms + Microsoft Lists + SharePoint;
Seleção de colaboradores diretamente das listas corporativas;
Menor dependência de digitação manual;
Fonte única de informações;
Controle centralizado;
Identificação automática de conflitos;
Melhor experiência para analistas;
Estrutura mais escalável e sustentável.
💡 Principais Melhorias da Versão 2.1
Integração completa com Microsoft Lists;
Centralização das informações no SharePoint;
Eliminação da busca manual por IPNs;
Utilização da identidade corporativa dos usuários;
Melhor qualidade dos dados recebidos;
Redução significativa de erros operacionais;
Identificação automática de conflitos;
Criação do Centro de Manutenção do Sistema;
Recuperação da área operacional;
Melhor experiência para analistas e supervisores;
Arquitetura mais moderna e integrada ao Microsoft 365.
📂 Estrutura do Projeto

hour-extra-management/

├── README.md

├── IMAGES/

├── PowerQuery/

│ ├── Colaboradores

│ ├── DisponibilidadeHoraExtra

│ └── SolicitaçõesHoraExtra

└── VBA/

├── EstaPastaDeTrabalho

├── FrmProgress

├── Hora-Extra

└── Modulo1

📸 Repositório de Evidências

O diretório IMAGES contém capturas de tela documentando todo o fluxo operacional, incluindo:

Integração com SharePoint;
Microsoft Forms;
Microsoft Lists;
Disponibilidade de Hora Extra;
Solicitações;
Importação de atividades;
Identificação de conflitos;
Resumo consolidado de horas;
Recuperação da área operacional;
Verificação do sistema;
Componentes VBA e Power Query.
👨‍💻 Autor

Gabriel Pereira

Projeto desenvolvido como iniciativa de melhoria de processo, automação operacional e integração de soluções Microsoft 365.

A solução foi criada com foco na modernização do controle de Hora Extra, redução de atividades manuais, melhoria da qualidade dos dados e ampliação da visibilidade operacional para analistas e supervisores.