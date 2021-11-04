Imports System.Text
Imports uPLibrary.Networking.M2Mqtt
Imports uPLibrary.Networking.M2Mqtt.Messages

Public Class FrmPrincipal

    Dim criptografiaEscolhida As Integer = -1
    Dim modoEscolhido As Integer = -1
    Dim chaveSalva As String

    Public lerDados As String

    '' MQTT
    Dim client As MqttClient
    Dim msgMqtt As StringBuilder = New StringBuilder(4096)

    '' AO INICIAR:
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Pegar nome usuario:
        'TextBox2.Text = My.User.Name


        'cifraVigenereCodificar("Teste", "ed123")
        Dim mensagem As String = "In playfair"

        'MsgBox(TripleDESEncrypt(mensagem, True))
        'MsgBox(TripleDESDecrypt(TripleDESEncrypt(mensagem, True), True))


        ' Verifica se usuário tem rede:
        If My.Computer.Network.IsAvailable Then
            'Dim strHostName = System.Net.Dns.GetHostName()
            'Dim strIPAddress = System.Net.Dns.Resolve(strHostName).AddressList(0).ToString()

            Dim ipLocal As System.Net.IPAddress = GetLocalIP()

            'TextBox1.Text = ipLocal.ToString
            'TextBox1.Text = "127.0.0.1"

        Else
            MsgBox("Esta desconectado!?")
        End If

    End Sub

    '' FUNCAO PEGAR IP
    Function GetLocalIP() As System.Net.IPAddress
        Dim localIP() As System.Net.IPAddress = System.Net.Dns.GetHostAddresses(System.Net.Dns.GetHostName)
        For Each current As System.Net.IPAddress In localIP
            If current.ToString Like "*[.]*[.]*[.]*" Then
                Try
                    Dim SplitVar() As String = current.ToString.Split(".")
                    If Len(SplitVar(0)) <= 3 And Len(SplitVar(1)) <= 3 And Len(SplitVar(2)) <= 3 And Len(SplitVar(3)) <= 3 Then
                        Return current
                    End If
                Catch ex As Exception
                End Try
            End If
        Next
    End Function

    '' BOTAO LIGAR
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If (TextBox1.Text.Length <> 0) Then
            Try
                client = New MqttClient(TextBox1.Text)

                Dim clientId As String = Guid.NewGuid().ToString()

                AddHandler client.MqttMsgPublishReceived, AddressOf Client_MqttMsgPublishReceived
                AddHandler client.ConnectionClosed, AddressOf Client_Disconnect

                client.Connect(clientId)

                If client.IsConnected Then
                    mostrarMensagem("# CONECTADO AO BROKER: " & TextBox1.Text)
                    ToolStripStatusLabel1.Text = "CONECTADO AO BROKER: '" & TextBox1.Text & "'"
                Else
                    mostrarMensagem("# DESCONECTADO")
                    ToolStripStatusLabel1.Text = "DESCONECTADO"
                End If

            Catch ex As Exception
                mostrarMensagem("# ERRO: " & ex.Message)
                ToolStripStatusLabel1.Text = "ERRO"
                MsgBox(ex.Message(), MsgBoxStyle.Critical)
            End Try
        Else
            ToolStripStatusLabel1.Text = "Please enter a valid Broker address "
        End If
    End Sub

    '' FUNCOES MQTT - Desconectou
    Private Sub Client_Disconnect(sender As Object, e As EventArgs)
        mostrarMensagem("# DESCONECTADO: ")
        ToolStripStatusLabel1.Text = "DESCONECTADO"
    End Sub

    '' FUNCOES MQTT - Chegou mensagem
    Private Sub Client_MqttMsgPublishReceived(ByVal sender As Object, ByVal e As MqttMsgPublishEventArgs)
        'msgMqtt.Append("[" + TimeOfDay.ToString("hh:mm:ss") + "] Topic: " + e.Topic.ToString() + ", Len: " + e.Message.Length.ToString() + ", Qos: " + e.QosLevel.ToString())
        msgMqtt.Append("[" + TimeOfDay.ToString("hh:mm:ss") & "] Topic: " & e.Topic.ToString() & ", Qos: " & e.QosLevel.ToString() & " ")
        msgMqtt.Append("Msg: " + Encoding.Default.GetString(e.Message))
        mostrarMensagem(msgMqtt.ToString)
        mostrarMensagem("DECRYPT: " & desCriptografarMensagem(criptografiaEscolhida, modoEscolhido, Encoding.Default.GetString(e.Message), chaveSalva))

    End Sub

    ' Ao selecionar o tipo de criptografia:
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

        ' Limpa todos os modos:
        ComboBox2.Items.Clear()

        'Desliga a seleção de modos(modo unico)
        ComboBox2.Enabled = False

        criptografiaEscolhida = ComboBox1.SelectedIndex
        modoEscolhido = ComboBox2.SelectedIndex
        chaveSalva = TextBox4.Text

        modoEscolhido = -1

        ' Conforme o tipo de criptografia:
        If ComboBox1.SelectedIndex < 6 Then ' Nada :)
            'Cifras
        ElseIf ComboBox1.SelectedIndex = 7 Or ComboBox1.SelectedIndex = 8 Or ComboBox1.SelectedIndex = 9 Then ' Se for DES, 3DES ou AES:
            ComboBox2.Enabled = True
            ComboBox2.Items.AddRange(New Object() {"ECB", "CBC", "CFB", "OFB", "CTR"})
            ComboBox2.Focus()
        Else 'Caso seja outro tipo:
            'MsgBox("Em desenvolvimento")
            'ComboBox1.SelectedIndex = 0
        End If

        '0 Sem Criptografia
        '1 Chave Simétrica
        '2 Cifra de Cesar
        '3 Cifra de Vigenère
        '4 Cifra One - time pad
        '5 Cifra Playflair
        '6 Cifra de Hill
        '7 DES
        '8 3DES
        '9 AES


        '' Conforme o tipo de criptografia:
        'If ComboBox1.SelectedIndex = 0 Then 'Chave Simétrica
        'ElseIf ComboBox1.SelectedIndex = 1 Then 'Cifra de Cesar
        'ElseIf ComboBox1.SelectedIndex = 2 Then 'Cifra de Vigenère
        'ElseIf ComboBox1.SelectedIndex = 3 Then 'Cifra One - time pad
        'ElseIf ComboBox1.SelectedIndex = 4 Then 'Cifra Playflair
        'ElseIf ComboBox1.SelectedIndex = 4 Then 'Cifra de Hill
        'ElseIf ComboBox1.SelectedIndex = 5 Then 'DES
        '    ComboBox2.Enabled = True
        '    ComboBox2.Items.AddRange(New Object() {"ECB", "CBC", "CFB", "OFB", "CTR"})
        '    ComboBox2.Focus()

        'ElseIf ComboBox1.SelectedIndex = 6 Then '3DES
        '    ComboBox2.Enabled = True
        '    ComboBox2.Items.AddRange(New Object() {"ECB", "CBC", "CFB", "OFB", "CTR"})
        '    ComboBox2.Focus()

        'ElseIf ComboBox1.SelectedIndex = 7 Then 'AES
        '    ComboBox2.Enabled = True
        '    ComboBox2.Items.AddRange(New Object() {"ECB", "CBC", "CFB", "OFB", "CTR"})
        '    ComboBox2.Focus()

        'ElseIf ComboBox1.SelectedIndex = 8 Then 'IDEA
        '    MsgBox("Em desenvolvimento")
        '    ComboBox1.SelectedIndex = 0
        'ElseIf ComboBox1.SelectedIndex = 9 Then 'SAFER
        '    MsgBox("Em desenvolvimento")
        '    ComboBox1.SelectedIndex = 0
        'ElseIf ComboBox1.SelectedIndex = 10 Then 'Chave Assimétrica
        '    MsgBox("Em desenvolvimento")
        '    ComboBox1.SelectedIndex = 0
        'ElseIf ComboBox1.SelectedIndex = 11 Then 'DESX
        '    MsgBox("Em desenvolvimento")
        '    ComboBox1.SelectedIndex = 0
        'ElseIf ComboBox1.SelectedIndex = 12 Then 'Camellia
        '    MsgBox("Em desenvolvimento")
        '    ComboBox1.SelectedIndex = 0
        'ElseIf ComboBox1.SelectedIndex = 13 Then 'RSA
        '    MsgBox("Em desenvolvimento")
        '    ComboBox1.SelectedIndex = 0
        'ElseIf ComboBox1.SelectedIndex = 14 Then 'Blowfish
        '    MsgBox("Em desenvolvimento")
        '    ComboBox1.SelectedIndex = 0
        'ElseIf ComboBox1.SelectedIndex = 15 Then 'Twofish
        '    MsgBox("Em desenvolvimento")
        '    ComboBox1.SelectedIndex = 0
        'ElseIf ComboBox1.SelectedIndex = 16 Then 'RC
        '    MsgBox("Em desenvolvimento")
        '    ComboBox1.SelectedIndex = 0
        'ElseIf ComboBox1.SelectedIndex = 17 Then 'ElGamal
        '    MsgBox("Em desenvolvimento")
        '    ComboBox1.SelectedIndex = 0
        'Else
        '    MsgBox("Erro, em desenvolvimento")
        '    ComboBox1.SelectedIndex = 0
        'End If



    End Sub

    ' Escolher o modo:
    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        modoEscolhido = ComboBox2.SelectedIndex
    End Sub

    '' MOSTRA MENSAGEM NA TELA:
    Public Sub mostrarMensagem(texto As String)
        lerDados = texto
        mensagemInterno()
    End Sub
    Private Sub mensagemInterno()
        If Me.InvokeRequired Then
            Me.Invoke(New MethodInvoker(AddressOf mensagemInterno))
        Else
            RichTextBox1.AppendText(lerDados & vbNewLine)
            'RichTextBox1.Update()

            RichTextBox1.SelectionStart = RichTextBox1.TextLength
            RichTextBox1.ScrollToCaret()
        End If
    End Sub

    '' BOTAO SUB:
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        If (client IsNot Nothing AndAlso client.IsConnected()) Then
            If (TextBox3.Text.Length <> 0) Then
                Try
                    Dim Topic() As String = {TextBox3.Text}
                    Dim Qos() As Byte = {0} 'At most once (0) At least once (1) Exactly once(2)
                    client.Subscribe(Topic, Qos)
                    mostrarMensagem("# INSCRITO EM: " & TextBox3.Text)
                    Button3.Enabled = True 'Liga botão enviar

                Catch ex As Exception
                    mostrarMensagem("# ERRO: " & ex.Message)
                    ToolStripStatusLabel1.Text = "ERRO"
                    MsgBox(ex.Message, MsgBoxStyle.Critical)

                End Try
            Else
                mostrarMensagem("# INFORME UM TOPICO VALIDO ")
                'ToolStripStatusLabel1.Text = "Please enter a valid topic "
            End If
        Else
            mostrarMensagem("# DESCONECTADO, FALHOU")
            'ToolStripStatusLabel1.Text = "Disconnected !! subscription procedure is not valid"
        End If

    End Sub

    '' BOTAO ENVIAR MENSAGEM
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        criptografiaEscolhida = ComboBox1.SelectedIndex
        modoEscolhido = ComboBox2.SelectedIndex
        chaveSalva = TextBox4.Text

        Try
            If TextBox5.Text.Length > 1 Then
                'mostrarMensagem("ENVIANDO: " & TextBox4.Text)
                Dim mensagemCriptografada As String = criptografarMensagem(ComboBox1.SelectedIndex, ComboBox2.SelectedIndex, TextBox5.Text, TextBox4.Text)

                Dim cronometro As New Stopwatch

                Try
                    cronometro.Start()
                    If (client IsNot Nothing AndAlso client.IsConnected()) Then

                        Try
                            Dim Qos As Byte = 0
                            client.Publish(TextBox3.Text, Encoding.Default.GetBytes(mensagemCriptografada), Qos, False)

                            mostrarMensagem("# PUBLICANDO: " & TextBox5.Text)
                            mostrarMensagem("# CRIPTOGRAFIA: " & mensagemCriptografada)

                            'ToolStripStatusLabel1.Text = "Publish to " + "{" + TextBox4.Text + "}"
                        Catch ex As Exception
                            mostrarMensagem("# ERRO AO PUBLICAR MSNG: " & ex.Message)
                            'ToolStripStatusLabel1.Text = "Error"
                            MsgBox(ex.Message, MsgBoxStyle.Critical)
                        End Try

                        'mostrarMensagem("# INFORM ")
                        'ToolStripStatusLabel1.Text = "Please enter a valid topic "

                    Else
                        mostrarMensagem("# DESCONECTADO, FALHOU")
                        'ToolStripStatusLabel1.Text = "Disconnected !! Publish procedure is not valid"

                        cronometro.Stop()
                    End If
                    'mostrarMensagem("ENVIOU: " & TextBox4.Text & " (" & cronometro.ElapsedMilliseconds.ToString & "ms)")
                    'ToolStripStatusLabel2.Text = "Tempo de resposta: " & cronometro.ElapsedMilliseconds.ToString & "ms"

                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try

            Else

                My.Computer.Audio.PlaySystemSound(Media.SystemSounds.Asterisk)
            End If

            TextBox5.Text = ""
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Function criptografarMensagem(tipo As Integer, modo As Integer, mensagem As String, chave As String) As String
        If tipo = 0 Then 'Sem Criptografia
            Return mensagem
        ElseIf tipo = 1 Then 'Chave Simétrica

        ElseIf tipo = 2 Then 'Cifra de Cesar
            MsgBox("OBS.: Precisa ser só numeros! (0~4)")
            Return cifraCesarCodificar(mensagem, chave)
        ElseIf tipo = 3 Then 'Cifra de Vigenère 'SO ACEITA LETRAS!
            MsgBox("OBS.: Chave Precisa conter só letras")
            Return cifraVigenereCodificar(mensagem, chave)
        ElseIf tipo = 4 Then 'Cifra One - time pad 'tem que ter tamanhos iguals
            MsgBox("OBS.: Chave deve ter o mesmo tamanho da mensagem")
            Return cifraOTPCodificar(mensagem, chave)
        ElseIf tipo = 5 Then 'Cifra Playflair 'Usar palavras pequenas
            MsgBox("OBS.: Em caso de erros, utilize mensagens menores")
            Return cifraPlayfairCodificar(mensagem, chave)
        ElseIf tipo = 6 Then 'Cifra de Hill     'Enviar {1,2,3,4}
            MsgBox("OBS.: Enviar uma Matriz em forma de vetor: '1,2,3,4'")
            Return cifraVigenereCodificar(mensagem, "teste")
        ElseIf tipo = 7 Then 'DES
            Return DESEncrypt(mensagem, chave)
        ElseIf tipo = 8 Then '3DES
            Return TripleDESEncrypt(mensagem, True)
        ElseIf tipo = 9 Then 'AES
            Return AES_Encrypt(mensagem, chave)
        Else
            MsgBox("Em desenvolvimento.")
        End If

        Return mensagem
    End Function

    Private Function desCriptografarMensagem(tipo As Integer, modo As Integer, mensagem As String, chave As String) As String
        If tipo = 0 Then 'Sem Criptografia
            Return mensagem
        ElseIf tipo = 1 Then 'Chave Simétrica

        ElseIf tipo = 2 Then 'Cifra de Cesar
            Return cifraCesarDecodificar(mensagem, chave)
        ElseIf tipo = 3 Then 'Cifra de Vigenère 'SO ACEITA LETRAS!
            Return cifraVigenereDecodificar(mensagem, chave)
        ElseIf tipo = 4 Then 'Cifra One - time pad 'tem que ter tamanhos iguals
            Return cifraOTPDecodificar(mensagem, chave)
        ElseIf tipo = 5 Then 'Cifra Playflair 'Usar palavras pequenas
            Return cifraPlayfairDecodificar(mensagem, chave)
        ElseIf tipo = 6 Then 'Cifra de Hill     'Enviar {1,2,3,4}
            Return cifraVigenereDecodificar(mensagem, "teste")
        ElseIf tipo = 7 Then 'DES
            Return DESDecrypt(mensagem, chave)
        ElseIf tipo = 8 Then '3DES
            Return TripleDESDecrypt(mensagem, True)
        ElseIf tipo = 9 Then 'AES
            Return AES_Decrypt(mensagem, chave)
        Else
            MsgBox("Em desenvolvimento.")
        End If

        Return mensagem
    End Function

    '' AO DAR ENTER NA MENSAGEM
    Private Sub TextBox5_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox5.KeyDown
        If e.KeyData = Keys.Enter Then
            Button3.PerformClick()
        End If
    End Sub

End Class