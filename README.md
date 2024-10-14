**AUTOMATIZAÇÃO DE PESQUISAS EM BASE DE DADOS - ARQUIVOS EM EXCEL**
Utilização de Regex; Loops, bibliotecas
versão final com esboço de interface gráfica



Explicação Detalhada do Código 

**1 - Importações Necessárias:**
tkinter: Biblioteca padrão para interfaces gráficas em Python.
filedialog e messagebox: Módulos do tkinter para seleção de arquivos e exibição de mensagens.
ttk: Módulo do tkinter para widgets temáticos, como Treeview.
threading: Para executar a pesquisa em uma thread separada, evitando que a interface trave durante a execução.
Outras bibliotecas como openpyxl, re, datetime e time são usadas para a lógica de pesquisa.

**2 - Funções Principais:**
converter_celula_para_string(celula): Converte o conteúdo de uma célula para string, tratando datas e valores None.
extrair_mprj(concatenated_row): Utiliza uma expressão regular para extrair o MPRJ da linha concatenada.
pesquisar_em_aba(sheet, pattern, abas_encontradas, resultados_tree, nome_pesquisa): Pesquisa em cada aba do Excel e insere os resultados na Treeview.
realizar_pesquisa(filename, nomes, resultados_tree): Carrega o workbook e executa a pesquisa para cada nome.
iniciar_pesquisa(filename, nomes, resultados_tree, btn_pesquisar): Inicia a pesquisa em uma thread separada para não bloquear a interface e gerencia o estado do botão de pesquisa.
selecionar_arquivo(entry_arquivo): Abre uma janela para seleção do arquivo Excel e atualiza o campo de entrada com o caminho do arquivo selecionado.
obter_nomes(input_nomes): Processa a string de entrada para obter uma lista de nomes a serem pesquisados.
salvar_resultados(resultados_tree): Salva os resultados exibidos na Treeview em um arquivo CSV.
criar_interface(): Cria a interface gráfica completa utilizando tkinter.

**3 - Interface Gráfica:**
Seleção do Arquivo Excel: Um LabelFrame contendo um Entry para mostrar o caminho do arquivo e um botão "Selecionar" para abrir o diálogo de seleção.
Entrada de Termos de Pesquisa: Outro LabelFrame com um Entry para inserir os nomes a serem pesquisados, separados por ; ou ,.
Botões de Ação: Botões para "Pesquisar", "Salvar Resultados" e "Limpar Resultados".
Área de Resultados: Um Treeview configurado com colunas para exibir os resultados da pesquisa, juntamente com uma barra de rolagem vertical.

**4 - Execução em Thread Separada:**
A função iniciar_pesquisa inicia a pesquisa em uma thread separada usando o módulo threading. Isso garante que a interface gráfica permaneça responsiva durante a execução da pesquisa, evitando que o aplicativo trave ou congele.

**5 - Salvar Resultados:**
A função salvar_resultados permite que você salve os resultados exibidos na Treeview em um arquivo CSV para referência futura.

**6 - Tratamento de Exceções:**
As funções possuem tratamento de exceções para capturar e exibir erros que possam ocorrer durante a execução, garantindo que o usuário seja informado adequadamente em caso de falhas.

**COMO UTILIZAR O PROGRAMA**
**A - Executar o Script:**

Salve o código acima em um arquivo Python, por exemplo, pesquisa_gui.py.

Certifique-se de ter as bibliotecas necessárias instaladas. Você pode instalar o openpyxl se ainda não tiver:

pip install openpyxl

Execute o script:
python pesquisa_gui.py

**B - Selecionar o Arquivo Excel:**

Clique no botão "Selecionar" para abrir uma janela de diálogo e escolha o arquivo Excel que deseja pesquisar.

**C - Inserir os Termos de Pesquisa:**

No campo "Nomes", insira os nomes que deseja pesquisar, separados por ; ou ,. Por exemplo:

Copiar código
João; Maria, Pedro

**D - Iniciar a Pesquisa:**
Clique no botão "Pesquisar" para iniciar a busca. A pesquisa será executada em uma thread separada, e os resultados serão exibidos na área de resultados conforme forem encontrados.
Visualizar os Resultados:

Os resultados aparecerão na Treeview com as colunas: Nome, Aba, Linha, MPRJ, Trecho e Linha Completa.

**E - Salvar Resultados:**
Após a pesquisa, você pode salvar os resultados clicando no botão "Salvar Resultados". Será solicitado que você escolha o local e o nome do arquivo CSV onde os resultados serão salvos.

**F - Limpar Resultados:**
Se desejar limpar os resultados exibidos na Treeview, clique no botão "Limpar Resultados".

Melhorias e Personalizações Adicionais
I. Progress Bar:
Para melhorar a experiência do usuário, você pode adicionar uma barra de progresso que indique o andamento da pesquisa. Isso pode ser feito utilizando o widget ttk.Progressbar e atualizando-o conforme a pesquisa avança.

II. Filtros Avançados:
Implementar opções para filtrar os resultados com base em critérios adicionais, como data ou outros parâmetros.

III. Exportar para Outros Formatos:
Além de CSV, você pode adicionar opções para exportar os resultados em outros formatos, como Excel ou PDF.

IV. Interface Mais Intuitiva:
Melhorar a disposição dos elementos da interface para torná-la mais intuitiva e agradável visualmente, utilizando temas do ttk ou outras bibliotecas de estilização.

V. Validação de Entrada:
Adicionar validações para garantir que o usuário selecione um arquivo válido e insira termos de pesquisa adequados antes de iniciar a busca.

**Considerações Finais
Este script fornece uma interface gráfica completa para realizar pesquisas em lote em arquivos Excel, tornando o processo mais acessível e amigável. Ele mantém toda a funcionalidade do seu script original e adiciona a facilidade de uso proporcionada por uma interface gráfica.**
