'                             Online VB Compiler.
'                 Code, Compile, Run and Debug VB program online.
' Write your code in this editor and press "Run" button to execute it.
Sub media()

DIM nota as integer
DIM nota1 as integer
DIM nota2 as integer
DIM nota3 as integer
DIM WS AS WORKSHEET

Application.screenupdate=false
Application.calculation=xl calculationmanual

SET WS = THISWORKBOOK.SHEETS("MEDIA")

nota =inputbox("digite sua 1 nota: ")				
nota1=inputbox("digite sua 2 nota: ")
nota2=inputbox("digite sua 3 nota: ")
nota4=inputbox("digite sua 4 nota: ")

media=(nota+nota1+nota2+nota3)/4

MSG("MEDIA")


IF MEDIA >= 7 THEN Interior.Color = RGB(VERDE)
Else 
Interior.Color = RGB(255, 0, 0) ' Vermelho
End if

application.calculation=xlcalculationautomatic
application.screenupdate=true

END SUB
