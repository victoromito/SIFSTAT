SIFSTAT - Sistema Unificado de Conferência de Mapas Estatísticos JBS
=======================================================================

O QUE É
-------
Este é o SIFSTAT unificado: um único programa com menu lateral que reúne
os três conferidores que você já usava separadamente:
  - Produção
  - Recebimento
  - Comercialização

A lógica de cálculo de cada módulo (leitura de ERP/PGA, localização de
colunas, tolerância de 0,0001 kg, cadastro de países da Comercialização
etc.) foi mantida exatamente como nos programas originais. Só a interface
foi unificada em um único menu.

COMO USAR (rodando o .py diretamente)
--------------------------------------
1. Instale o Python 3 (se ainda não tiver).
2. Instale as dependências:
       pip install -r requirements.txt
3. Rode o programa:
       python sifstat.py

COMO GERAR O .EXE (Windows, com PyInstaller)
----------------------------------------------
1. pip install pyinstaller
2. Coloque um ícone chamado "icone.ico" dentro da pasta "assets" (opcional).
3. Rode, a partir da pasta do projeto:

   pyinstaller --onefile --windowed --name SIFSTAT ^
       --add-data "assets;assets" ^
       sifstat.py

   O executável final aparecerá em dist\SIFSTAT.exe
   (o parâmetro --add-data garante que a logo da JBS e o ícone
   sejam embutidos no .exe).

ESTRUTURA DE PASTAS
--------------------
sifstat_unificado/
  sifstat.py            -> programa principal (todo o sistema)
  assets/
    jbs_logo.png         -> logo usada no rodapé do menu lateral
    icone.ico             -> (opcional) ícone do programa/instalador
    tutorial/              -> capturas de tela usadas na aba "Como utilizar ?"
  requirements.txt
  LEIA-ME.txt

A aba "Como utilizar ?" já está com o passo a passo completo (exportação do
ERP, exportação do PGA-SIGSIF e comparação no SIFSTAT), com rolagem e as
imagens de exemplo. Se precisar trocar alguma imagem do tutorial, é só
substituir o arquivo correspondente dentro de assets/tutorial mantendo o
mesmo nome.

OBSERVAÇÕES
------------
- A logo da JBS incluída em assets/jbs_logo.png foi recortada a partir do
  print de tela enviado, como um placeholder. Se vocês tiverem o arquivo
  oficial da logo (PNG com fundo transparente, em boa resolução), é só
  substituir esse arquivo pelo oficial - o restante do programa não precisa
  mudar.
- A página "Como utilizar ?" já está funcional (é possível clicar e abrir),
  mas está propositalmente sem conteúdo ainda, conforme solicitado.
- A planilha "Config_Paises_SIFStat.xlsx" (memória de países da
  Comercialização) continua sendo criada/lida ao lado do executável,
  exatamente como no programa original.
