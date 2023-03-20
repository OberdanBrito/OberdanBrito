<img src="https://github.com/OberdanBrito/OberdanBrito/blob/95409f8d18dbcc1be009fed6e4e0f9108b03fe38/profile-pic.png" width="128"/>

😁 Olá, pessoal, desde já, agradeço por lê a minha apresentação, sinta-se a vontade para entrar em contato quando quiser! Ou convidar para um café/cerveja 🍻.

Gosto de desenvolver códigos: há muito tempo faço isso. Até onde, me lembro bem... em 1994, eu digitava em teclados duros da IBM e olhava para monitores com prompt verde, e esquecia de colocar o disquete de 5- 1/4 antes de enviar o comando para liberar o disquete corrente. E sempre gostei de descobrir o funcionamento das coisas.

Boa parte das coisas que faço aqui no GitHub estão no privado, mas de uns tempos para cá, resolvi que vou publicar algumas coisas que faço semanalmente. Geralmente partes de códigos que compõe sistemas. O intuito aqui é apenas ajudar.


Tenho experiência em diversas áreas, incluindo desenvolvimento web, aplicativos móveis e análise de dados. Alguns dos meus projetos mais recentes incluem uma plataforma que auxilia condôminos para conveniência em suas moradias. Mas além desse, consumo desenvolver do zero, códigos que solucionam em diferentes áreas, indústria, escolas, escritórios etc.


#### **Essa lista a baixo já mostra logo o que mais gosto né?**

[![Top Langs](https://github-readme-stats.vercel.app/api/top-langs/?username=oberdanbrito)](https://github.com/oberdanbrito/github-readme-stats)



Vamos lá!
## **2023-03-20 Uma ajudinha com Ms Excel**

O código dessa semana é uma ajuda que dei a um amigo. Ele estava precisando separar uma lista de clientes contida em um arquivo, onde continha a palavra "Empresa" porém estava separada por intervalos com nomes de funcionários.
Como a urgência falava mais alto, construir algo do zero seria impossível, então encontrei uma solução simples e que acabou sendo útil para muitas planilhas dele.
[Este e o link do código completo](https://gist.github.com/OberdanBrito/253fc530539c3e72d6268826829151be.js)

Há sim... quem não se lembra do velho VBA. Ainda na década de 90 a Microsoft precisava apresentar aos seus clientes corporativos alguma forma que ajudasse eles a automatizar as tarefas. Visto que a grande sacada do Office é oferecer um produto genérico na qual usuários com conhecimento mais aprofundados pudessem deixar rotinas mais inteligentes.
mas sem longas histórias, a solução que encontrei, utiliza dois loops para identificar onde encontrar uma palavra que sempre repete no arquivo. Se essa for a sua necessidade dê uma olhada nesse exemplo:

Para quem não está familiarizado com VBA, toda variável deve ser declarada e repare que para fazer isso você deve usar a palavra reservada "Dim" de dimensionar, sacou?

```
 
    Dim flag As Boolean 
    flag = False 
    
    Dim linha, contador, inicio, final As Long
    linha = 1 
    contador = 0 
    inicio = 0 
    final = 0 

```  

Agora a parte fundamental, repare que há um loop. 
Este é utilizado para percorrer todas as linhas da planilha do Excel.

```
    While Not flag 
       If InStr(ws.Cells(linha, 1).Value, "Empresa:") > 0 Then

```
Você pode substituir por qualquer palavra para pesquisar, desde que essa faça parte de um padrão dentro do seu arquivo. Vamos imaginar que ao invés de "Empresa:" o seu arquivo seja uma lista de alunos, nesse caso basta modificar o valor da pesquisa pela palavra "alunos".
Mas atenção, procure identificar bem o seu padrão. Se na sua planilha existir mais de uma forma para escrita, você deve primeiro certificar-se de que está pesquisando um caso bem específico.
No meu caso para evitar essa coincidência, eu reparei que sempre quando havia a palavra empresa ela era seguida pelos dois pontos (:), assim ficou fácil.

Após a identificação, você deve estabelecer um ponto de partida e um ponto de encerramento, que servem para você fazer o que mais estiver necessitado no momento, uma cópia das células (Meu caso), formatação ou edição de dados seja possível.
Então para que a magia pudesse ocorrer eu usei um novo loop


````
    Do
        final = final + 1
        If InStr(ws.Cells(final, 1).Value, "Empresa:") > 0 Or vazios = 10 Then Exit Do
        If IsEmpty(ws.Cells(final, 1)) Then vazios = vazios + 1
        
    Loop
    Range("A" & inicio & ":K" & final - 1).Copy
````

Repare que no primeiro loop estamos percorrendo linha por linha para identificar onde começa uma empresa, já nesse segundo caso nós precisamos identificar o final que determina o início da uma outra empresa.
Uma vez encontrado o final agora sabemos o que selecionar. É aí que entra a palavra "Range"
Essa função interna do Excel nada mais é que a capacidade de selecionar uma área ou os mesmos movimentos que você faria com o seu mouse passando encima e selecionado as áreas que deseja de uma planilha

Por fim eu definir o que desejava com a minha seleção. **É claro que você deve alterar isso também, a fim de refletir a sua necessidade 👀 **

Bônus: O código final apresenta uma maneira fácil de separar o conteúdo obtido no comando "Range" e cola dentro de um novo arquivo.
Um para cada empresa que foi encontrada na planilha.


```visual badic
    Workbooks.Add
    Set novoarquivo = ActiveWorkbook
    Worksheets.Item(1).Name = "Planilha da fatura"
    Worksheets.Item(1).Paste
    Selection.PasteSpecial Paste:=xlPasteColumnWidths, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Range("A1").Activate
    
    novoarquivo.SaveAs Filename:="C:\minhas_empresas\" & empresa & ".xls"
    novoarquivo.Save
    novoarquivo.Close
    
    Debug.Print empresa & " Inicio:" & inicio & " Final:" & final
```

Pessoal, acessem o [código completo](https://gist.github.com/OberdanBrito/253fc530539c3e72d6268826829151be)
estudem e se divirtam! Se tiver alguma dúvida entrem em contato.

Uma boa semana e até a próxima. 



[![wakatime](https://wakatime.com/badge/user/eb9c14f3-847b-4b7f-be05-24cba40f2b44.svg)](https://wakatime.com/@eb9c14f3-847b-4b7f-be05-24cba40f2b44)

![Snake animation](https://github.com/oberdanbrito/oberdanbrito/blob/output/github-contribution-grid-snake.svg)
