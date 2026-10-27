# Networking da Direita: formulário de cadastro

Página com a bandeira do Brasil animada. Cada cadastro vira uma linha numa planilha.

Há duas formas de usar:

| | Online (Vercel + Google Planilhas) | Só no computador |
|---|---|---|
| Quem acessa | Qualquer pessoa com o link | Só o seu PC e celulares no mesmo Wi-Fi |
| Onde caem os cadastros | Planilha do Google | `dados/cadastros-networking.xlsx` |
| Precisa do PC ligado? | Não | Sim |

---

## Online: Vercel + Google Planilhas

### 1. Criar a planilha e o script (5 min)

1. Acesse **sheets.new** e dê o nome **Networking da Direita – Cadastros**.
2. Menu **Extensões → Apps Script**.
3. Apague o que aparecer, cole todo o conteúdo de `google-apps-script/Codigo.gs` e clique em **Salvar**.
4. No topo, escolha a função **setup** e clique em **Executar**. Na primeira vez o Google pede autorização: **Revisar permissões → sua conta → Avançado → Acessar (não seguro) → Permitir**. O aviso aparece porque o script é seu. É normal.
5. Volte à planilha: a aba **CADASTROS** estará criada e formatada.

### 2. Publicar o script (2 min)

1. No Apps Script: **Implantar → Nova implantação**. Na engrenagem, escolha **App da Web**.
2. **Executar como:** Eu. **Quem pode acessar:** Qualquer pessoa.
3. Clique em **Implantar** e copie a **URL do App da Web** (termina em `/exec`).
4. Teste: abra essa URL no navegador. Deve aparecer `"Formulário do Networking da Direita ativo."`

### 3. Ligar o site à planilha (1 min)

1. No GitHub, abra `site/config.js` → ícone de lápis (**Edit**).
2. Cole a URL entre as aspas:
   ```js
   window.PLANILHA_URL = "https://script.google.com/macros/s/XXXXXXXX/exec";
   ```
3. **Commit changes**. A Vercel publica sozinha em 1 ou 2 minutos.
4. Faça um cadastro de teste no site. Ele aparece na planilha como **Novo** (amarelo).

> Se mudar o `Codigo.gs` depois, use **Implantar → Gerenciar implantações → lápis → Versão: Nova versão**. Assim a URL continua a mesma.

Opcional: em `CONFIG.EMAIL_AVISO`, no `Codigo.gs`, coloque um e-mail para receber aviso a cada cadastro.

---

## Só no computador

Dê **dois cliques em `iniciar.command`** e deixe a janela do Terminal aberta. O formulário abre no navegador, e os celulares no mesmo Wi-Fi usam o endereço que aparece no Terminal.

Os cadastros caem em `dados/cadastros-networking.xlsx`, com cópia em `dados/backup-cadastros.csv`. Feche a planilha no Excel enquanto o formulário recebe cadastros.

Na primeira vez, se o Mac disser "não pode ser aberto": botão direito em `iniciar.command` → **Abrir**. Requisitos: Python 3 e `openpyxl` (`pip3 install openpyxl`).

---

## Lista de profissionais na página

Logo abaixo do formulário fica **Profissionais da nossa comunidade**:
- A lista vem da planilha e é separada por categoria.
- Tem busca por nome, profissão, empresa, cidade ou serviço, e filtros de categoria e cidade.
- Cada cartão traz os botões **Visitar Instagram** e **Entrar em contato**. O botão de contato só aparece se a pessoa informou o WhatsApp.
- Um cadastro novo aparece na lista na hora, sem recarregar a página.
- A página busca a lista de novo a cada minuto e sempre que a pessoa volta para a aba.

**Lista só para quem se cadastrou:**
- Quem ainda não se cadastrou vê a seção borrada, com cadeado, o número de profissionais e o botão "Fazer meu cadastro". O servidor não envia os dados dessas pessoas para quem não tem cadastro.
- Ao se cadastrar, a pessoa recebe uma chave de acesso, que fica guardada no navegador, e a lista abre na hora.
- Em outro celular ou navegador, basta clicar em "Já me cadastrei" e digitar o Instagram do cadastro.
- Quem estiver com Status Recusado ou Inativo perde o acesso.
- Limite: quem souber o Instagram de alguém cadastrado consegue liberar a lista por "Já me cadastrei". Para um guia de comunidade, isso é um meio-termo razoável.

**Moderação:**
- Por padrão, o cadastro aparece no site assim que é enviado.
- Para tirar alguém do ar, mude o Status para **Recusado** ou **Inativo**. Some do site em até 30 segundos.
- Para exigir aprovação antes de publicar, mude `EXIGIR_APROVACAO` para `true`. Fica no `CONFIG` do `Codigo.gs` (Google) ou no início do `servidor.py` (versão local). Com isso, só aparece quem estiver com Status **Aprovado**.
- O mesmo Instagram não pode ser cadastrado duas vezes.

**Depois de mudar o `Codigo.gs`**, publique a nova versão: **Implantar → Gerenciar implantações → lápis → Versão: Nova versão → Implantar**. A URL continua a mesma.

## A planilha (as duas formas usam as mesmas colunas)

- **Status:** chega como **Novo**. Mude para Aprovado, Recusado ou Inativo pela lista suspensa.
- **Link do Instagram:** é clicável.
- **Observações da administração:** livre para anotações.
- Use as setinhas do cabeçalho para filtrar, por exemplo só os "Novo" ou uma categoria.

## Arquivos

| Arquivo | Para que serve |
|---|---|
| `site/` | A página publicada (a Vercel usa esta pasta, conforme o `vercel.json`) |
| `site/config.js` | Onde vai a URL da planilha do Google |
| `google-apps-script/Codigo.gs` | Script que grava na planilha do Google |
| `servidor.py`, `iniciar.command` | Versão só no computador |
| `src/` | Fonte da página. Depois de editar `src/pagina.html`, rode `python3 src/montar.py` |
