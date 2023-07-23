# **Projeto VB6 CRUD com DataGrid**

## Descrição

Este projeto VB6 é uma aplicação simples que permite realizar operações CRUD (Create, Read, Update, Delete) em um banco de dados. Ele carrega os valores de um banco em um DataGrid, permite adicionar e excluir valores do banco, atualiza o DataGrid após as operações e também permite imprimir os valores do DataGrid em um arquivo PDF.

## Sobre o VB6

O Visual Basic 6 (VB6) é uma linguagem de programação de alto nível e um ambiente de desenvolvimento integrado (IDE) criado pela Microsoft. Embora seja uma tecnologia mais antiga, ainda é amplamente utilizado em projetos legados e em ambientes onde a migração para versões mais recentes do Visual Basic não é viável. O VB6 foi conhecido por sua simplicidade e facilidade de uso, permitindo que os desenvolvedores criassem rapidamente aplicativos para Windows.

## Funcionalidades

1.  **Carregar valores do banco em um DataGrid**: Ao iniciar a aplicação, os valores armazenados no banco de dados serão carregados e exibidos em um DataGrid na interface do usuário.
    
2.  **Adicionar valores ao banco**: O usuário pode inserir novos valores na aplicação, que serão adicionados ao banco de dados e atualizados no DataGrid.
    
3.  **Excluir valores do banco**: O usuário pode buscar por um registro específico pelo ID e, em seguida, excluí-lo do banco. O DataGrid será atualizado automaticamente após a exclusão.
    
4.  **Imprimir valores do DataGrid em PDF**: O usuário pode imprimir os valores exibidos no DataGrid em um arquivo PDF, facilitando a geração de relatórios.

## Código do Projeto

O código fornecido demonstra duas principais funcionalidades do CRUD: busca de registro pelo ID e exclusão do registro. Essas funcionalidades estão associadas a três botões:

1.  **btnBuscarExcluir_Click()**: Esse botão busca os valores do banco com base no ID fornecido pelo usuário e exibe as informações nos campos de texto correspondentes (txtNomeExcluir, txtEmailExcluir, txtTelefoneExcluir e txtIdadeExcluir).
    
2.  **btnExcluir_Click()**: Esse botão executa a exclusão do registro que foi buscado anteriormente pelo ID fornecido. A exclusão é realizada diretamente no banco de dados, e o DataGrid é atualizado para refletir as alterações.
    
3.  **btnFecharExcluir_Click()**: Esse botão fecha a janela de exclusão quando o usuário deseja sair da funcionalidade de exclusão.

## Tutorial de uso

1.  Baixe os arquivos do projeto do GitHub: `demo_vb6.exe` e `Db.mdb` (o arquivo do banco de dados).
2.  Crie uma pasta chamada `VB6_Demo` na unidade `C:` do computador.
3.  Coloque os arquivos baixados (`demo_vb6.exe` e `Db.mdb`) dentro da pasta `VB6_Demo`.
4.  Execute o arquivo `demo_vb6.exe` para iniciar a aplicação.
5.  A aplicação será aberta, e os valores do banco de dados serão carregados no DataGrid automaticamente.
6.  Use os botões fornecidos para adicionar, excluir e atualizar os valores no DataGrid.
7.  Para gerar um arquivo PDF com os valores do DataGrid, use a funcionalidade de impressão disponível na aplicação.

Aviso: Certifique-se de que o ambiente em que você está executando a aplicação tenha suporte ao VB6. Se necessário, instale o VB6 Runtime e outras dependências para garantir a compatibilidade do projeto.

Observação: Este projeto é apenas uma demonstração de funcionalidades básicas e não representa uma aplicação completa e robusta.

