<h1>Programa da Expedição</h1> 
O projeto tem o objetivo de automatizar a criação do relatório da expedição de um pedido. Neste relatório, estão presentes as informações sobre as dimensões de cada volume, os produtos que eles contém, seu peso individual e o peso total. Para fazer a seleção de quais produtos estarão em quais volumes, são utilizadas diversas condicionais a partir das medidas de um item e seu tipo.

<h2>Organização dos diretórios</h2>
<h3>/database</h3>
    <ul>
    <li>db_config</li>
    </ul>
<h3>/src</h3>
    <ul>
    <li>/classes</li>
    <ul>produto.py</ul>
    <ul>relatório.py</ul>
    <li>/images</li>
    <li>/modules</li>
    <li>main.py</li>
    </ul>


<h2>Tecnologias utilizadas</h2>
A principal linguagem utilizada para a construção deste projeto foi Python e para a conexão com o banco de dados foram utilizados comandos SQL. 
<h3>As principais bibliotecas utilizadas foram:</h3>
<ul>
    <li>Tkinter para a composição gráfica;</li>
    <li>Pandas e numpy para a organização dos dados e inserção em um documento xlsx;</li>
    <li>OS para as ações que dependem de funcionalidades do sistema operacional;</li>
    
</ul>
<h2>Etapas do Projeto</h2>
    <ul>
        <li>Obter os dados corretos do Sysop;</li>
        <li>Adicionar informações em objetos de produtos;</li>
        <li>Inserir todos os produtos em uma tabela;</li>
        <li>Organizar os volumes na tabela;</li>
        <li>Calcular os pesos;</li>
        <li>Criar estrutura gráfica;</li>
        <li>Exibir relatório no centro da página;</li>
        <li>Criar Tela para alterações no relatório;</li>
        <li>Criar questionário para validação final;</li>
    </ul>
<h2>Estrutura visual</h2>
A parte visual do projeto é divida em duas páginas: a página principal e a de configurações.
<h3>Página principal</h3>
No cabeçalho da página, é onde está presente o campo para a inserção do pedido. No corpo da página é onde será visualizado o resutado final do relatório por meio de uma treeview. E a barra lateral é onde estão presentes as informações gerais do pedido, além do botão que realiza a impressão do arquivo.
<h3>Página de configurações</h3>
Aqui é onde o usuário poderá fazer algumas alterações gerais no projeto. A página é dividida em três:
área onde será feita a inserção de novos registros no banco de dados que contém os valores de peso, área para alterar os valores dos registros e a área para alterar algumas condições que são feitas na criação do projeto.

<h2>Conexão com banco de dados</h2>
Informações necessárias do banco de dados a partir do número do pedido:
    - Ambiente;
    - Largura;
    - Altura;

<h2>Criação do relatório</h2>
    - Obter o número do pedido;
    - Verificar se existe algum arquivo com o número do pedido;
    - Caso exista, exibir, caso contrário criar um novo.
