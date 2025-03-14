Public Class frmAsistNuevProm
    '************************************************************************************
    'Importante: Se requiere que el formulario donde se ubica este código, contenga un
    'TextBox y un ListBox con los nombres "TextBox1" y "ListBox1".
    '************************************************************************************

    'Lista donde se cargarán las palabras del autocompletado
    Dim sugerencias As New AutoCompleteStringCollection()
    Private Function LeerPalabraActual(textBox As Windows.Forms.TextBox) As String
        ' Obtiene la posición actual del cursor
        Dim cursorPos As Integer = textBox.SelectionStart
        ' Si el cursor está al principio del texto, no hay palabra actual
        If cursorPos = 0 Then Return Nothing
        ' Obtener el texto antes del cursor
        Dim textoAntesDelCursor As String = textBox.Text.Substring(0, cursorPos)
        ' Buscar la última palabra antes del cursor
        Dim palabraInicio As Integer = textoAntesDelCursor.LastIndexOf(" "c) + 1
        ' Extraer la palabra actual
        Dim palabraActual As String = textoAntesDelCursor.Substring(palabraInicio)
        Return palabraActual
    End Function
    Private Function ObtenerPosicionCursorMultilinea(textBox As Windows.Forms.TextBox) As Point
        ' Obtener la posición del cursor
        '**** Aparece defasado
        Dim cursorPos As Integer = textBox.SelectionStart

        ' Obtener la línea actual y la columna dentro de esa línea
        Dim lineaActual As Integer = textBox.GetLineFromCharIndex(cursorPos)
        Dim columnaActual As Integer = cursorPos - textBox.GetFirstCharIndexFromLine(lineaActual)

        ' Obtener el texto de la línea actual
        Dim textoLineaActual As String = textBox.Lines(lineaActual)

        ' Calcular el ancho del texto desde el inicio de la línea hasta la columna actual
        Using g As Graphics = textBox.CreateGraphics()
            Dim tamanoTexto As SizeF = g.MeasureString(textoLineaActual.Substring(0, columnaActual), textBox.Font)

            ' Calcular la posición del cursor
            Dim posicionCursor As New Point(CInt(tamanoTexto.Width), lineaActual * textBox.Font.Height)
            Return textBox.PointToScreen(posicionCursor)
        End Using
    End Function
    Private Sub PosicionarListBoxCursor(textBox As Windows.Forms.TextBox, listBox As ListBox)
        Dim pos = ObtenerPosicionCursorMultilinea(textBox)
        pos.Y += textBox.Font.Height    'Baja al nivel dela sig. línea
        listBox.Location = Me.PointToClient(pos)   'Corrige coordenadas
    End Sub
    Private Sub ReemplazarPalabraBajoCursor(newWord As String)
        ' Obtener la posición del cursor
        Dim cursorPosition As Integer = TextBox1.SelectionStart

        ' Obtener el texto antes y después del cursor
        Dim textBeforeCursor As String = TextBox1.Text.Substring(0, cursorPosition)
        Dim textAfterCursor As String = TextBox1.Text.Substring(cursorPosition)
        ' Buscar el inicio de la palabra antes del cursor
        Dim startOfWord As Integer = textBeforeCursor.LastIndexOf(" "c) + 1
        ' Buscar el final de la palabra después del cursor
        Dim endOfWord As Integer = textAfterCursor.IndexOf(" "c)
        If endOfWord = -1 Then
            endOfWord = textAfterCursor.Length
        End If
        ' Obtener la palabra completa
        Dim word As String = TextBox1.Text.Substring(startOfWord, cursorPosition - startOfWord + endOfWord)

        ' Reemplazar la palabra con el nuevo valor
        Dim newText As String = TextBox1.Text.Substring(0, startOfWord) & newWord & TextBox1.Text.Substring(cursorPosition + endOfWord)

        ' Actualizar el texto del TextBox con la palabra reemplazada
        TextBox1.Text = newText
        ' Colocar el cursor al final de la palabra reemplazada
        TextBox1.SelectionStart = startOfWord + newWord.Length
    End Sub
    Private Sub InicAutocompletado()
        ' Crear una lista de sugerencias
        sugerencias.AddRange(New String() {"Por la compra de ",
            "Uva", "Aguacate", "Manzana", "Maracuya", "Banana", "Cereza", "Dátil", "Fresa", "Granada"})
        ListBox1.Visible = False
        ListBox1.SelectionMode = SelectionMode.One
    End Sub
    Private Function BuscarListaSugerencias() As List(Of String)
        ' Filtrar las sugerencias basadas en lo que el usuario escribe
        Dim palabraAct As String = LeerPalabraActual(TextBox1)
        If palabraAct = "" Then
            ListBox1.Visible = False
            Return New List(Of String)   'Devuelve arreglo vacío
        End If
        Return sugerencias.Cast(Of String)().Where(Function(item) item.StartsWith(palabraAct, StringComparison.InvariantCultureIgnoreCase)).ToList()
    End Function
    Private Sub procesarAutocompletado()
        'Lee la palabra actual en el caudro de texto y muestra, si aplica, la lista de
        'autocompletado en la posición apropiada.
        Dim filteredList As List(Of String) = BuscarListaSugerencias()
        If filteredList.Count > 0 Then
            ListBox1.Items.Clear()
            'Agrega datos al ListBox
            ListBox1.Items.AddRange(filteredList.ToArray())
            'Posiciona y muestra la lista sobre el cursor
            PosicionarListBoxCursor(TextBox1, ListBox1)
            ListBox1.Visible = True
            ListBox1.SelectedIndex = 0
        Else
            ListBox1.Visible = False
        End If
    End Sub
    Private Sub txtDescripcion_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        procesarAutocompletado()
    End Sub
    Private Sub txtDescripcion_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter And ListBox1.Visible And ListBox1.SelectedIndex >= 0 Then
            'Hay sugerencia seleccionada y se pide elegir una
            Dim selIndex As Integer = ListBox1.SelectedIndex
            Dim sug_seleccionada As String = ListBox1.SelectedItem.ToString()
            ReemplazarPalabraBajoCursor(sug_seleccionada)
            e.SuppressKeyPress = True
            Exit Sub
        End If
        'Procesa la tecla
        procesarAutocompletado()
        If e.KeyCode = Keys.Escape Then
            ListBox1.Visible = False
        ElseIf e.KeyCode = Keys.Down Then
            If ListBox1.Visible And ListBox1.SelectedIndex >= 0 Then
                'Hay lista de autocompletado visible
                If ListBox1.SelectedIndex > ListBox1.Items.Count - 2 Then
                    Exit Sub
                End If
                ListBox1.SelectedIndex = ListBox1.SelectedIndex + 1
                e.SuppressKeyPress = True
            Else

            End If
        ElseIf e.KeyCode = Keys.Tab Then
        End If
    End Sub
    Private Sub txtDescripcion_Leave(sender As Object, e As EventArgs) Handles TextBox1.Leave
        ListBox1.Visible = False ' Ocultar el ListBox cuando el foco se pierde
    End Sub

    Public Sub mostrar()
        'Muestra el formulario actual e inicia el autocompletado.
        InicAutocompletado()
        Me.Show()
    End Sub
End Class