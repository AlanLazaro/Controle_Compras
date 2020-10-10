Sub ConferindoPedido()

	linha_Nome = 1

	linha_fim_SP = Sheets("-").Range("A1").End(xlDown).Row
	linha_fim_Nome = Sheets("Nomes").Range("A1").End(xlDown).Row

	While linha_Nome <= linha_fim_Nome
		nome = Sheets("Nomes").Cells(linha_Nome, 1)
		linha_SP = 2
		
		While linha_SP <= linha_fim_SP
		
			If nome = Sheets("-").Cells(linha_SP, 2) Then
				Sheets("-").Cells(linha_SP, 2).Select
				
				With Selection.Interior
					.Pattern = xlSolid
					.PatternColorIndex = xlAutomatic
					.Color = 255
					.TintAndShade = 0
					.PatternTintAndShade = 0
				End With
				
			End If
			
			linha_SP = linha_SP + 1
		Wend
		
		linha_Nome = linha_Nome + 1
	Wend

End Sub
