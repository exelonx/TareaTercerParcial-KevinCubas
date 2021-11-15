Imports System.IO
Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'Valida espacio en blanco
        If busqueda.Text = Nothing Then
            Exit Sub
        End If

        Try
            Dim lector As New StreamReader(TareaTercerParcial_KevinCubas.My.Computer.FileSystem.CurrentDirectory & "\tabla.txt")
            Dim lineaTexto As String = ""
            Dim fila As Integer = 0         'Indica en que fila se encuentra la palabra buscada
            Dim columna As Integer = 0      'Indica en que columna se encuentra la palabra buscada
            Dim match As Integer = 0        'Cuenta cada vez que haya un match de letras
            Dim lineaIteracion As Integer = 0

            'limpiar datagrid y textbox de filas y columnas
            DataGridView1.ClearSelection()
            filaPalabra.Clear()
            columnaPalabra.Clear()

            While lector.Peek > 0
                lineaTexto = lector.ReadLine
                match = 0 'Reiniciar el contador para validar 
                For i = 0 To lineaTexto.ToArray.Length - 1    'Recorre cada letra de la linea de texto
                    'Letra de la tabla = Letra de la palabra a buscar
                    If lineaTexto.Substring(i, 1).ToLower() = LTrim(busqueda.Text.Substring(match, 1).ToLower()) Then
                        match += 1  'Incrementa el numero de matches (se usa en la comparación de letras)
                        'Comprueba si el número de matches es igual al número de letras de la palabra a buscar
                        If busqueda.Text.Length = match Then
                            fila = lineaIteracion       'Recibe el número de fila
                            columna = (i + 1) - match   'Calcula la posición de columna

                            'Despliega las filas y columnas
                            filaPalabra.Text = fila
                            columnaPalabra.Text = columna

                            'Marca la palabra en el DataGridView
                            For j = columna To columna + match - 1
                                DataGridView1.Rows(fila).Cells(j).Selected = True
                            Next
                            lector.Close()          'Cierra el streamReader si el match se completa
                            Exit Sub                'Cierra el ciclo si el match se completa
                        End If
                    Else
                        match = 0       'Reinicia el número de matches para que vuelva a comparar con la primera letra de la palabra a buscar
                    End If
                Next
                lineaIteracion += 1
            End While

            'En caso de no encontrar un match completo
            lector.Close()              'Cierra el streamReader si el match NO se completa
            MsgBox("Palabra no encontrada")
        Catch ex As Exception
            MsgBox("Error " & vbCrLf & ex.Message)
        End Try

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Despliega el archivo en el datagrid
        Try
            Dim lector As New StreamReader(TareaTercerParcial_KevinCubas.My.Computer.FileSystem.CurrentDirectory & "\tabla.txt")
            Dim lineaTexto As String = ""
            Dim linea As Integer = 0        'Contador para iterar

            While lector.Peek > 0
                DataGridView1.Rows.Add()
                lineaTexto = lector.ReadLine
                For columna = 0 To lineaTexto.ToArray.Length - 1  'Recorrera la distancia de la cadena
                    'agrega la letra que esta en esta posición al datagrid
                    DataGridView1.Rows(linea).Cells(columna).Value = lineaTexto.Substring(columna, 1)
                Next
                linea += 1  'Pasa a la siguiente fila
            End While
            lector.Close()
        Catch ex As Exception
            MsgBox("Error " & vbCrLf & ex.Message)
        End Try

        DataGridView1.ClearSelection()

    End Sub
End Class
