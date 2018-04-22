'Observação: 
'-Criar evento com tag auxilar de gatilho com scan de 5000
'-Cria um xc, adicionar uma duas propriedades: tipo variante para receber o endereço do objeto, tipo integer para selecionar função


on error Resume next
Dim max
Dim min
Dim avg

min = 999999
avg = 0
max = 0


'Selecionar maior valor	 de temperatura entre as 22 Vav's
If xc_vav_max.seletor =  1 Then 
	For each VAV in xc_vav_max.Teste	
		If VAV.Temp.Quality = 192  then
			If VAV.Temp.Value > max then
				max = VAV.Temp.Value
			Else
				max = max
			End If
		Else
			max = max
		End If
	Value = max
	Next
End if


'Selecionar menor valor	 de temperatura entre as 22 Vav's
If xc_vav_max.seletor =  2 Then
	For each VAV in xc_vav_max.Teste
		If VAV.Temp.Quality = 192  then
			If VAV.Temp.Value < min then
				min = VAV.Temp.Value
			Else
				min = min
			End If
		Else
			min = min
		End If	
	Next
	Value = min
End if


'Selecionar valor medio	de temperatura entre as 22 Vav's
Dim i
i = 0
If xc_vav_max.seletor =  3 Then
	For each VAV in xc_vav_max.Teste
		If VAV.Temp.Quality = 192  then
			avg = avg+VAV.Temp.Value
			i = i+1
		End If
	Next
Value = avg/i
End if


