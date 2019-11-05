// BIBLIOTECAS NECESSÁRIAS
#Include "TOTVS.ch

//------------------------------------------------------
// ENVIA E-MAIL DE DESCONEXÃO
//-------------------------------------------------------
Main Function NewsMail()
    Local cMail     := Decode64("Z3VpbGhlcm1lLmJpZ29pc0B0b3R2cy5jb20uYnI=") // E-MAIL PARA ENVIO
    Local cPass     := Decode64("aTA0ZzAyYjE5OTZId3Rm")                     // SENHA DO E-MAIL
    Local cSMTP     := "smtp.gmail.com"                                     // SMTP DO E-MAIL
    Local cLog      := Space(0)                                             // LOGS DE ERRO
    Local nPort     := 0                                                    // PORTA DE COMUNICAÇÃO
    Local nSecurity := 2                                                    // TIPO DE SEGURANÇA
    Local nTimeout  := 60                                                   // TEMPO DE REQUISIÇÃO
    Local xRet      := .T.                                                  // RETORNO/VALIDADOR
    Local oServer   := TMailManager():New()                                 // CLIENTE DO SERVIDOR DE E-MAIL
    Local oMessage  := TMailMessage():New()                                 // CONSTRUTOR DA MENSAGEM

    // DEFINE A SEGURANÇA DA COMUNICAÇÃO COM O SMTP
    If (nSecurity == 0)
        nPort := 25 // PORTA PADRÃO DE COMUNICAÇÃO DO SMTP
    ElseIf (nSecurity == 1)
        nPort := 465 // PORTA PARA COMUNICAÇÃO DO SMTP COM SSL
        oServer:SetUseSSL(.T.)
    Else
        nPort := 587 // PORTA PARA COMUNICAÇÃO DO SMTP COM TLS
        oServer:SetUseTLS(.T.)
    EndIf

    BEGIN SEQUENCE
        // INICIA A CONEXÃO COM O SERVIDOR SMTP
        xRet := oServer:Init(Space(0), cSMTP, cMail, cPass, NIL, nPort)
        If (!xRet == 0)
            cLog := NoAcento("Não foi possível estabelecer a conexão com o servidor: ") + oServer:GetErrorString(xRet)
            BREAK
        EndIf

        // DEFINE O TEMPO MÁXIMO DA REQUISIÇÃO
        xRet := oServer:SetSMTPTimeout(nTimeout)
        If (!xRet == 0)
            cLog := NoAcento("Não foi possível definir o time-out de ") + CValToChar(nTimeout) + " segundos"
            BREAK
        EndIf

        // ESTABELECE A CONEXÃO COM O SERVIDOR SMTP
        xRet := oServer:SMTPConnect()
        If (!xRet == 0)
            cLog := NoAcento("Não foi possível estabelecer conexao com o servidor de SMTP: ") + oServer:GetErrorString(xRet)
            BREAK
        EndIf

        // REALIZA A AUTENTICAÇÃO NO SERVIDOR
        xRet := oServer:SMTPAuth(cMail, cPass)
        If (!xRet == 0)
            cLog := NoAcento("Não foi possível se autenticar no servidor de SMTP: ") + oServer:GetErrorString(xRet)
            oServer:SMTPDisconnect()
            BREAK
        EndIf

        // FORMATA O CORPO DA NOVA MENSAGEM
        oMessage:Clear()
        oMessage:cDate    := CValToChar(Date())
        oMessage:cFrom    := cMail
        oMessage:cTo      := "guilherme.bigois@totvs.com.br"
        oMessage:cSubject := "TOTVS Protheus | Prime News"
        oMessage:cBody    := GetMailBody()

        // REALIZA O ENVIO DO E-MAIL
        xRet := oMessage:Send(oServer)
        If (!xRet == 0)
            cLog := NoAcento("Não foi possível realizar o envio da mensagem: ") + oServer:GetErrorString(xRet)
            BREAK
        EndIf

        // DESCONECTA DO SERVIDOR SMTP
        xRet := oServer:SMTPDisconnect()
        If (!xRet == 0)
            cLog := NoAcento("Não foi possível realizar a desconexão do servidor SMTP: ") + oServer:GetErrorString( xRet )
            BREAK
        EndIf

        // MENSAGEM DE SUCESSO CASO NÃO TENHA GERADO EXCESSÃO
        ConOut(NoAcento("E-mail enviado como sucesso!"))
    RECOVER
        Final(cLog) // GERA EXCESSÃO EM CASO DE ERRO
    END SEQUENCE
Return (NIL)

//------------------------------------------------------
// RETORNA O CORPO DO E-MAIL A SER ENVIADO
//------------------------------------------------------
Static Function GetMailBody()
    Local cHTML := NRead("top.html") + NRead("banner.html") + NRead("header.html")
    Local cBreak := NRead("break.html")

    cHTML += NCreateNews(CToD("13/02/2019", "DDMMYYYY"), "EUA pressiona Google, Facebook e Twitter", "https://g1.globo.com/economia/tecnologia/noticia/2019/10/29/ue-pressiona-google-facebook-e-twitter-a-agirem-mais-contra-desinformacao.ghtml") + cBreak
    cHTML += NCreateNews(CToD("27/05/2019", "DDMMYYYY"), "Lucro da Alphabet, dona da Google, cai 23%", "https://g1.globo.com/economia/noticia/2019/10/28/lucro-da-alphabet-dona-do-google-cai-23percent-no-3o-tri-para-us-7-bilhoes.ghtml") + cBreak
    cHTML += NCreateNews(CToD("29/10/2019", "DDMMYYYY"), "Falha em chip de celular permite rastrear aparelho", "https://g1.globo.com/economia/tecnologia/blog/altieres-rohr/post/2019/10/28/falha-em-chip-de-celular-permite-rastrear-localizacao-do-aparelho.ghtml")
    // cHTML += NRead("valid.html")

    cHTML += NRead("footer.html")
Return (DecodeUTF8(cHTML))

Static Function NRead(cFile)
    Local cHTML := Space(0)
    Local oFile := FwFileReader():New("\dirdoc\mail\" + cFile)

    If (oFile:Open())
        cHTML := oFile:FullRead()
    Else
        Final("File " + cFile + " not found!")
    EndIf
Return (cHTML)

Static Function NCreateNews(dDate, cTitle, cLink)
    Local cHTML := '<table border="0" cellpadding="0" cellspacing="0" width="100%" class="mcnTextBlock" style="min-width: 100%;border-collapse: collapse;mso-table-lspace: 0pt;mso-table-rspace: 0pt;-ms-text-size-adjust: 100%;-webkit-text-size-adjust: 100%;"><tbody class="mcnTextBlockOuter"><tr><td valign="top" class="mcnTextBlockInner" style="padding-top: 9px;mso-line-height-rule: exactly;-ms-text-size-adjust: 100%;-webkit-text-size-adjust: 100%;"><table align="left" border="0" cellpadding="0" cellspacing="0" style="max-width: 100%;min-width: 100%;border-collapse: collapse;mso-table-lspace: 0pt;mso-table-rspace: 0pt;-ms-text-size-adjust: 100%;-webkit-text-size-adjust: 100%;" width="100%" class="mcnTextContentContainer"><tbody><tr><td valign="top" class="mcnTextContent" style="padding: 0px 18px 9px;font-size: 18px;text-align: center;mso-line-height-rule: exactly;-ms-text-size-adjust: 100%;-webkit-text-size-adjust: 100%;word-break: break-word;color: #7c7c7c;font-family:' + "'Open Sans', 'Helvetica Neue', Helvetica, Arial, sans-serif;line-height: 150%;" + '">'

    cHTML += '<span style="font-size:13px">' + DToC(dDate, "DDMMYYYY") + '</span><br>'
    cHTML += '<span style="color:#000000">' + cTitle + '</span><br>'
    cHTML += '<span style="font-size:12px"><a href="' + cLink + '"target="_blank" style="mso-line-height-rule: exactly;-ms-text-size-adjust: 100%;-webkit-text-size-adjust: 100%;color: #00a8bb;font-weight: normal;text-decoration: underline;">Ver Artigo</a></span></td></tr></tbody></table></td></tr></tbody></table>'
Return (cHTML)
