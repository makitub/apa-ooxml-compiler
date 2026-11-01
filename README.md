\# OOXML APA 7 Formatter \& Headless PDF Compiler



Um motor determinístico em Python concebido para inspeção, conformidade e normalização de documentos Microsoft Word (`.docx`) de acordo com as normas \*\*APA (7.ª edição)\*\*, integrando um compilador \*headless\* para exportação de PDF vetorial de alta fidelidade via automação COM.



\---



\## 📌 Contexto e Problema de Engenharia



Editores de texto visuais (WYSIWYG) como o Microsoft Word introduzem frequentemente anomalias de formatação decorrentes da disparidade entre a \*\*especificação subjacente dos dados\*\* e o \*\*motor de renderização gráfica\*\*:

1\. \*\*Derivação de Margens e Recuo Francês (\*Hanging Indent\*):\*\* Erros de arredondamento de escala de ecrã (\*DPI scaling\* / \*pixel-snapping\*) e herança de estilos residuais causam desalinhamentos visuais na régua e quebras incorretas de parágrafos.

2\. \*\*Perda de Fidelidade em PDF:\*\* Exportadores baseados em emulação de impressão virtual ou bibliotecas de terceiros muitas vezes desassociam tabulações, corrompem índices e quebram sumários dinâmicos.



Este projeto substitui a inferência probabilística e as correções manuais repetitivas por um \*\*pipeline determinístico\*\* que manipula diretamente os nós XML da norma \*\*Office Open XML (ECMA-376)\*\* em nível de unidades físicas (\*twips\* e \*dxa\*), compilando subsequentemente o artefacto final via API nativa de automação.



\---



\## 🚀 Funcionalidades Principais



\* \*\*Manipulação Nativa do OpenXML (OOXML):\*\* Modificação direta dos nós `<w:ind>`, `<w:spacing>` e `<w:rPr>` no `word/document.xml`.

\* \*\*Conformidade Estrita com a APA 7:\*\*

&#x20; \* Recuo pendente calibrado exatamente em `720 dxa` ($0,5''$ ou $1,27\\text{ cm}$).

&#x20; \* Espaçamento duplo uniforme (`480 twips`) sem espaçamento excedente entre blocos (`before=0`, `after=0`).

&#x20; \* Injeção de quebra de página mandatória (`<w:pageBreakBefore>`) na secção de referências.

&#x20; \* Parser cirúrgico de itálico: distingue obras autónomas (livros e normas técnicas) de verbetes de dicionários e artigos.

\* \*\*Compilação Headless de Alta Resolução (PDF):\*\*

&#x20; \* Inicialização silenciosa da instância COM do Microsoft Word (`pywin32`).

&#x20; \* Atualização forçada em lote de campos dinâmicos (`Fields.Update()`) e sumários (`TablesOfContents.Update()`).

&#x20; \* Exportação com fidelidade gráfica vetorial nativa a 600 DPI com suporte a marcadores internos.



\---



\## 🏗️ Arquitetura do Pipeline

