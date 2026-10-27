# Networking da Direita: formulário de cadastro (100% no seu computador)

Tudo roda na sua máquina:
- **Página:** o formulário com a bandeira animada.
- **Servidor:** recebe os envios.
- **Planilha:** onde ficam os cadastros.

Nada vai para Google, Netlify ou qualquer serviço de fora. Até as fontes estão na pasta.

## Ligar

Dê **dois cliques em `iniciar.command`**.

Abre uma janela do Terminal e o formulário aparece no navegador. **Deixe a janela do Terminal aberta** enquanto quiser receber cadastros. Para desligar, feche a janela ou aperte `Ctrl+C`.

Na primeira vez, o Mac pode mostrar:
- **"não pode ser aberto"**: clique com o botão direito em `iniciar.command` → **Abrir** → **Abrir**.
- **"aceitar conexões de entrada?"**: clique em **Permitir** para os celulares do Wi-Fi conseguirem acessar.

Pelo Terminal também funciona: `python3 servidor.py`

## Endereços

O Terminal mostra:

```
Neste computador:   http://localhost:8080
Celular no Wi-Fi:   http://192.168.x.x:8080
Planilha:           .../dados/cadastros-networking.xlsx
```

- **Neste computador:** abra no navegador do Mac.
- **Celular no Wi-Fi:** qualquer celular conectado **na mesma rede Wi-Fi** abre esse endereço. Bom para um evento presencial, por exemplo.
- Para funcionar só no seu Mac, sem acesso pelo Wi-Fi: `python3 servidor.py --so-local`

> Como o site está no seu computador, quem está fora da sua rede Wi-Fi (o pessoal do grupo em casa) não consegue abrir. Para isso seria preciso expor o computador à internet, o que vai contra a ideia de manter tudo local. O Mac também precisa estar ligado e com a janela aberta para receber cadastros.

## A planilha

Arquivo: `dados/cadastros-networking.xlsx` (criado sozinho na primeira vez que liga).

| Coluna | O que é |
|---|---|
| ID | NDD-0001, NDD-0002… (automático) |
| Data do cadastro | Data e hora do envio |
| Status | **Novo** (amarelo). Mude para Aprovado, Recusado ou Inativo pela lista suspensa |
| Nome completo … Interesse em parcerias | O que a pessoa preencheu |
| Link do Instagram | Clicável, abre o perfil |
| Autorizou divulgação | Sempre "Sim" (o formulário não envia sem autorização) |
| Observações da administração | Livre para anotações |

Use as setinhas do cabeçalho para filtrar (ex.: só os "Novo", ou uma categoria).

**Cuidado:** feche a planilha no Excel/Numbers enquanto o formulário estiver recebendo cadastros. Se ela estiver aberta e você salvar por cima, os cadastros que chegaram nesse meio-tempo podem sumir do .xlsx. Por segurança, **todo envio também vai para `dados/backup-cadastros.csv`**, então nada se perde.

## Arquivos da pasta

| Arquivo | Para que serve |
|---|---|
| `iniciar.command` | Liga tudo com dois cliques |
| `servidor.py` | Servidor local (recebe e grava os cadastros) |
| `site/` | A página do formulário, com fontes locais |
| `dados/` | Planilha e backup (aparece depois da primeira execução) |
| `src/` | Fonte da página. Se editar `src/pagina.html`, rode `python3 src/montar.py` para atualizar o `site/` |

Requisitos: Python 3 (já vem no Mac) e o pacote `openpyxl` (já instalado neste Mac; em outro computador use `pip3 install openpyxl`).
