' ============================================================
' 誤送信防止マクロ for Microsoft Outlook
' モジュール: ThisOutlookSession に貼り付けること
' 社内ドメイン: dfeii.or.jp
' ============================================================

Private Const INTERNAL_DOMAIN As String = "dfeii.or.jp"
Private Const LARGE_RECIPIENT_THRESHOLD As Long = 10

' 添付忘れを検出するキーワード一覧
Private Const ATTACH_KEYWORDS As String = "添付,別添,別紙,添付資料,ファイル,資料,参考資料,ご参照"

' ============================================================
' メイン処理: 送信イベントに自動で呼び出される
' ============================================================
Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)

    ' メールアイテム以外はスキップ
    If Item.Class <> olMail Then Exit Sub

    Dim bExternalFound  As Boolean
    Dim bInternalFound  As Boolean
    Dim sExternalList   As String
    Dim oRecip          As Recipient
    Dim sAddr           As String

    ' ── 1. 宛先の分類 ────────────────────────────────────
    For Each oRecip In Item.Recipients
        sAddr = GetSMTPAddress(oRecip)
        If sAddr = "" Then GoTo NextRecip

        If InStr(1, LCase(sAddr), "@" & LCase(INTERNAL_DOMAIN)) > 0 Then
            bInternalFound = True
        Else
            bExternalFound = True
            sExternalList = sExternalList & "  ・" & sAddr & vbCrLf
        End If
NextRecip:
    Next oRecip

    ' ── 2. 添付忘れチェック（完全ブロック） ──────────────
    If HasAttachmentKeyword(Item.Body) And Item.Attachments.Count = 0 Then
        MsgBox "【送信ブロック】添付ファイルがありません" & vbCrLf & vbCrLf & _
               "本文に添付を示すキーワードが含まれていますが、" & vbCrLf & _
               "ファイルが添付されていません。" & vbCrLf & vbCrLf & _
               "添付ファイルを確認してから再度送信してください。", _
               vbCritical, "誤送信防止 ― 送信ブロック"
        Cancel = True
        Exit Sub
    End If

    ' ── 3. 件名なしチェック（完全ブロック） ──────────────
    If Trim(Item.Subject) = "" Then
        MsgBox "【送信ブロック】件名が入力されていません" & vbCrLf & vbCrLf & _
               "件名を入力してから再度送信してください。", _
               vbCritical, "誤送信防止 ― 送信ブロック"
        Cancel = True
        Exit Sub
    End If

    ' ── 4. 社内・社外混在チェック（強い警告） ────────────
    If bInternalFound And bExternalFound Then
        Dim sMixMsg As String
        sMixMsg = "【警告】社内・社外が混在した宛先です！" & vbCrLf & vbCrLf & _
                  "外部宛先:" & vbCrLf & sExternalList & vbCrLf & _
                  "社内関係者と社外の方が同じメールに含まれています。" & vbCrLf & _
                  "送信内容を今一度ご確認ください。" & vbCrLf & vbCrLf & _
                  "このまま送信しますか？"

        If MsgBox(sMixMsg, vbExclamation + vbYesNo + vbDefaultButton2, _
                  "誤送信防止 ― 社内外混在 警告") = vbNo Then
            Cancel = True
            Exit Sub
        End If

    ' ── 5. 外部ドメインのみチェック（確認） ───────────────
    ElseIf bExternalFound Then
        Dim sExtMsg As String
        sExtMsg = "【確認】外部ドメイン宛てのメールです" & vbCrLf & vbCrLf & _
                  "送信先:" & vbCrLf & sExternalList & vbCrLf & _
                  "このまま送信しますか？"

        If MsgBox(sExtMsg, vbExclamation + vbYesNo + vbDefaultButton2, _
                  "誤送信防止 ― 外部送信 確認") = vbNo Then
            Cancel = True
            Exit Sub
        End If
    End If

    ' ── 6. 大量宛先チェック（警告） ───────────────────────
    If Item.Recipients.Count >= LARGE_RECIPIENT_THRESHOLD Then
        Dim sBulkMsg As String
        sBulkMsg = "【警告】送信先が " & Item.Recipients.Count & " 件あります" & vbCrLf & vbCrLf & _
                   "大量送信になります。宛先リストを今一度確認してください。" & vbCrLf & vbCrLf & _
                   "このまま送信しますか？"

        If MsgBox(sBulkMsg, vbExclamation + vbYesNo + vbDefaultButton2, _
                  "誤送信防止 ― 大量宛先 警告") = vbNo Then
            Cancel = True
            Exit Sub
        End If
    End If

End Sub

' ============================================================
' SMTP アドレスを取得するヘルパー関数
' Exchange アドレス（EX 形式）に対応
' ============================================================
Private Function GetSMTPAddress(oRecip As Recipient) As String
    On Error GoTo ErrorHandler

    Dim sAddr As String
    sAddr = ""

    ' Exchange アドレスの場合は ExchangeUser 経由で SMTP を取得
    If oRecip.AddressEntry.Type = "EX" Then
        Dim oExUser As ExchangeUser
        Set oExUser = oRecip.AddressEntry.GetExchangeUser()
        If Not oExUser Is Nothing Then
            sAddr = oExUser.PrimarySmtpAddress
        End If
    Else
        sAddr = oRecip.Address
    End If

    GetSMTPAddress = sAddr
    Exit Function

ErrorHandler:
    GetSMTPAddress = ""
End Function

' ============================================================
' 本文に添付を示すキーワードが含まれているか判定
' ============================================================
Private Function HasAttachmentKeyword(sBody As String) As Boolean
    Dim aKeywords() As String
    Dim keyword     As Variant

    aKeywords = Split(ATTACH_KEYWORDS, ",")

    For Each keyword In aKeywords
        If InStr(1, sBody, CStr(keyword)) > 0 Then
            HasAttachmentKeyword = True
            Exit Function
        End If
    Next keyword

    HasAttachmentKeyword = False
End Function
