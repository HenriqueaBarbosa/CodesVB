<h1>Planilha com auto-save</h1>

<p>1 - Para nossa planilha funcionar corretamente será preciso criar um módulo e um esta pasta de trabalho no excel aonde será colado nosso código em VB.NET</p>

>Pressione Alt+F11 para abrir o editor de Visual Basic (no Mac, pressione FN+ALT+F11) e clique em Inserir > Módulo. Uma nova janela de módulo aparece no lado direito do editor Visual Basic.

<p>2 - Depois da ciração do módulo basta copiar o código do do arquivo Módulo.vb do repositório e colar</p>

<p>3 - Para um melhor entendimento esse repositório foi separado por (módulo) e (esta pasta de trabalho)</p>

>Ainda no visual basic em VBAProject abra o modulo (esta pasta de trabalho) e cole o código do repositório com nome Esta-pasta-de-trabalho.vb

<h6>Pronto agora sua planilha estará com auto save a cada 60 segundos, que pode ser alterado de acordo com sua preferência no modulo na variável abaixo:
  
 ```VB.NET
 Public Const cRunIntervalSeconds = 60 '1 minuto 
