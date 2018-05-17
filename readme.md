# README NoteJam
- Baseado no aplicativo de [Sergey Komar][Komar];
- Versão feita por [Diego S. Silva][Silva]

## NoteJam VBScript + ASP + ADO Versão 1.4
- Aplicativo para gerenciar as suas anotações, escrevendo-as na web;
- Ele te ajuda a organizar suas anotações por pastas (pads), se desejar.

### Atualizações
- Agora os textos são exibidos em formato markdown;
- Conserto geral de bugs;
- Reestruturação do app conforme padrão MVC.

## Links para o NoteJam
- Visite [este link][link] para o projeto original;
- [Clique aqui][aqui] para poder acessar o sistema.

## Como Usar
- Simplesmente crie seu login em "Sign Up" com seu e-mail e crie uma senha;
	- Seu login para as próximas vezes será seu e-mail.
- Após isso, você poderá criar, editar e excluir suas anotações e as pastas
às quais pertence.

## Conteúdo

		/notejam-vbs-asp-ado
		|
		+-- asp
		|  |
		|  +- controller
		|  |   |
		|  |   +- login
		|  |   |  |
		|  |   |  +-------- forgot.asp
		|  |   |  +-------- login.asp
		|  |   |  +-------- update-password.asp
		|  |   +- notes
		|  |   |  |
		|  |   |  +-------- add-a-note.asp
		|  |   |  +-------- alter-a-note.asp
		|  |   |  +-------- kill-note.asp
		|  |   |  +-------- show-notes.asp
		|  |   |  +-------- single-note.asp
		|  |  +- pads
		|  |  |
		|  |  +-------- add-a-pad.asp
		|  |  +-------- alter-a-pad.asp
		|  |  +-------- kill-pad.asp
		|  |  +-------- show-pads.asp
		|  |  +-------- single-pad.asp
		|  |   
		|  +- model
		|     |
		|     +-------- model-login.asp
		|     +-------- model-notes.asp
		|     +-------- model-pads.asp
		|  
		+-- html
		|   |
		|   +---- account-settings.asp
		|   +---- alerts-and-errors.asp
		|   +---- create.asp
		|   +---- create-pad.asp
		|   +---- delete.asp
		|   +---- delete-pad.asp
		|   +---- edit-asp
		|   +---- empty-list.asp
		|   +---- forgot-password.asp
		|   +---- index.asp
		|   +---- pad-notes.asp
		|   +---- reset-password.asp
		|   +---- signin.asp
		|   +---- signup.asp
		|   +---- view.asp
		|   +- css
		|   |  |
		|   |  +-------- style.css
		|   +- screenshots
		|      |
		|      +- svgs
		|         |
		|         +-------- delete.svg
		|         +-------- edit.svg
		|
		+-- javascript
		|   |
		|   +---- moment.js
		|   +---- moments.asp
		|   +---- moment-with-locales.js
		|
		+-- sql
		|   |
		|   +---- notejam.sql
		|
		+-- readme.md

*****
[Komar]: https://github.com/komarserjio/
[link]: https://github.com/komarserjio/notejam
[Silva]: https://github.com/diegossilva-1995-01-25
[aqui]: https://blogs.lojcomm.com.br/diego.silva/notejam-vbs-asp-ado/html/signin.asp