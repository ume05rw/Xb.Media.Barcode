Option Strict On

''' <summary>
''' バーコード関連関数群
''' </summary>
''' <remarks>
''' ※注意※ 日本語文字列には対応出来ない。
''' </remarks>
Public Class Barcode


    ''' <summary>
    ''' バーコード区分
    ''' </summary>
    ''' <remarks>
    ''' バーコードの種類と特徴
    ''' http://www.media9.co.jp/m_tuhan/encyclopedia.html
    ''' 
    ''' なお、DotNetBarcodeで対応しているが Jan8, Jan13, Code39, Qr の4種のみ。
    ''' </remarks>
    Public Enum Type

        ''' <summary>
        ''' JANコード-13桁
        ''' </summary>
        ''' <remarks>
        ''' 文字種：半角数値
        ''' 桁数　：13桁固定
        ''' 流通コードとして、JISにより規格化されている。
        ''' ヨーロッパのEAN、アメリカのUPCと互換性がある。
        ''' 世界共通コード生活用品のほぼ全てに使用されている。
        ''' 食品、雑貨、書籍、雑誌、家電など。
        ''' </remarks>
        Jan13 = DotNetBarcode.Types.Jan13

        ''' <summary>
        ''' JANコード-8桁
        ''' </summary>
        ''' <remarks>
        ''' 文字種：半角数値
        ''' 桁数　：8桁固定
        ''' 流通コードとして、JISにより規格化されている。
        ''' ヨーロッパのEAN、アメリカのUPCと互換性がある。
        ''' 世界共通コード生活用品のほぼ全てに使用されている。
        ''' 食品、雑貨、書籍、雑誌、家電など。
        ''' </remarks>
        Jan8 = DotNetBarcode.Types.Jan8

        ''' <summary>
        ''' CODE39
        ''' </summary>
        ''' <remarks>
        ''' 文字種：半角数値,半角英字,記号["-", " "(スペース), "$", "/", "+", "%", "."],スタート/ストップコード["*"]。
        ''' 桁数　：可変
        ''' アルファベットや記号が扱えるため、品番などを表現できる。
        ''' 工業用バーコードとして多く使われている。日本電子機械工業会(EIAJ）
        ''' </remarks>
        Code39 = DotNetBarcode.Types.Code39

        '''' <summary>
        '''' CODE128
        '''' </summary>
        '''' <remarks>
        '''' 文字種：半角数値,半角英字(ASCIIコードは全てOK)
        '''' 桁数　：可変
        '''' 多くの文字が扱える。数字のみで表すなら、もっともサイズが小さくなる。（ただし12桁以上）
        '''' 米国UCCの混載用物流シンボル
        '''' </remarks>
        'Code128

        '''' <summary>
        '''' NW-7
        '''' </summary>
        '''' <remarks>
        '''' 文字種：半角数値,記号["-", "$", ":", "/", "+", "."],スタート/ストップコード["a", "b", "c", "d", "d" ]
        '''' 桁数　：可変
        '''' プリンタで印刷しやすいモジュール構成
        '''' 宅配便の伝票、DPE、図書館の貸し出し管理など。
        '''' </remarks>
        'Nw7

        '''' <summary>
        '''' ITF-14桁
        '''' </summary>
        '''' <remarks>
        '''' 文字種：半角数値
        '''' 桁数　：14桁固定
        '''' 高密度な印字が可能・物流商品コードとして1987年にJIS規格化されている。
        '''' 標準物流コードとしてＪＩＳ化し使われている。
        '''' </remarks>
        'Itf14

        '''' <summary>
        '''' ITF-16桁
        '''' </summary>
        '''' <remarks>
        '''' 文字種：半角数値
        '''' 桁数　：16桁固定
        '''' 高密度な印字が可能・物流商品コードとして1987年にJIS規格化されている。
        '''' 標準物流コードとしてＪＩＳ化し使われている。
        '''' </remarks>
        'Itf16

        ''' <summary>
        ''' QRコード
        ''' </summary>
        ''' <remarks>
        ''' 文字種：各種
        ''' 桁数　：可変、英数 1,523文字まで、漢字 644文字まで。
        ''' 代表的な２次元バーコード。日本デンソーで開発されたコードで2000年6月15日にISOの国際規格として制定された
        ''' 圧倒的な情報量が特長。
        ''' 日本自動車工業会が標準2次元コードとして採用するなど、各種分野で活用が進んでいる。
        ''' </remarks>
        Qr = DotNetBarcode.Types.QRCode

        '''' <summary>
        '''' CODE93
        '''' </summary>
        '''' <remarks>
        '''' 文字種：半角数値,記号["-", " "(スペース), "$", "/", "+", "%", "." ],半角英字大文字, スタート/ストップコード["□"], 制御文字4種
        '''' 桁数　：？
        '''' 制御キャラクタを用いることによりフルアスキーの表現が可能チェックデジットを2つ持っている
        '''' CODE39のバーコード長を短くする為に開発されたバーコードで、全てのASCIIコードを表すことが出来る。
        '''' </remarks>
        'Code93

        '''' <summary>
        '''' INDUSTRIAL2 OF 5
        '''' </summary>
        '''' <remarks>
        '''' 文字種：半角数値
        '''' 桁数　：可変
        '''' 物流管理などに使用されている。
        '''' 低密度でバー構成が単純であり、段ボール印刷に向いている。誤読率がすこし高い。
        '''' 元々工業用に多用されたが、数字しか表せず低密度のため CODE-39 などに置き換えられた。
        '''' </remarks>
        'Industrial2Of5

    End Enum


    ''' <summary>
    ''' 渡し値文字列でバーコード画像を生成する。
    ''' </summary>
    ''' <param name="value"></param>
    ''' <param name="type"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' DotNetBarcodeでバーコード画像を生成する。
    ''' </remarks>
    Public Shared Function GetImage(ByVal value As String, _
                                    ByVal type As Xb.Media.Barcode.Type, _
                                    ByVal width As Integer, _
                                    ByVal height As Integer) As System.Drawing.Image

        Dim graphics As System.Drawing.Graphics, _
            image As Drawing.Bitmap = New Drawing.Bitmap(width, height), _
            builder As DotNetBarcode = New DotNetBarcode()

        'PictureBoxからGraphicsオブジェクトを取得。
        graphics = Drawing.Graphics.FromImage(image)

        'Graphicsオブジェクトへバーコードを描画。
        builder.Type = CType(type, DotNetBarcode.Types)
        builder.WriteBar(value, 0, 0, width, height, graphics)

        '各オブジェクトを破棄。
        graphics.Dispose()
        builder = Nothing

        Return image

    End Function


    ''' <summary>
    ''' 画像ファイルからバーコードデータ文字列を取得する。
    ''' </summary>
    ''' <param name="fileName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetString(ByVal fileName As String) As String
        Dim currentPath, zbarPath, defaultArgs, xml As String, _
            doc As System.Xml.XmlDocument = New System.Xml.XmlDocument(), _
            nodes As System.Xml.XmlNodeList

        fileName = Xb.App.Path.GetAbsPath(fileName)
        If (Not IO.File.Exists(fileName)) Then
            Xb.Util.Out("Xb.Util.Barcode.GetString: 画像ファイルが見つかりません。")
            Throw New ArgumentException("画像ファイルが見つかりません。")
        End If

        currentPath = IO.Path.GetDirectoryName(Reflection.Assembly.GetExecutingAssembly().Location)
        zbarPath = "Lib\ZBar\zbarimg.exe"
        defaultArgs = "-D --xml"
        xml = Xb.App.Process.Execute(Xb.App.Path.GetAbsPath(zbarPath, currentPath), _
                                        defaultArgs & " " & Xb.Str.Dquote(fileName), _
                                        True)

        Try
            doc.LoadXml(xml)
            nodes = doc.DocumentElement.GetElementsByTagName("data")
        Catch ex As Exception
            Xb.Util.Out("Xb.Util.Barcode.GetString: バーコード解析に失敗しました。")
            Throw New ArgumentException("バーコード解析に失敗しました。")
        End Try

        If (nodes.Count <= 0) Then
            Xb.Util.Out("Xb.Util.Barcode.GetString: バーコードが検出出来ませんでした。")
            Throw New ArgumentException("バーコードが検出出来ませんでした。")
        End If

        Return nodes.Item(0).InnerText

    End Function


    ''' <summary>
    ''' 画像オブジェクトからバーコードデータ文字列を取得する。
    ''' </summary>
    ''' <param name="image"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' 画像オブジェクトを一旦ファイルに書き出したあと、ZBarに読ませる。
    ''' </remarks>
    Public Shared Function GetString(ByRef image As System.Drawing.Image) As String

        If (image Is Nothing) Then
            Xb.Util.Out("Xb.Util.Barcode.GetString: 画像オブジェクトが不正です。")
            Throw New ArgumentException("画像オブジェクトが不正です。")
        End If

        Dim tmpName, result As String

        tmpName = Xb.App.Path.GetTempFilename("png")
        image.Save(tmpName, System.Drawing.Imaging.ImageFormat.Png)

        Try
            result = GetString(tmpName)
        Catch ex As Exception
            IO.File.Delete(tmpName)
            Xb.Util.Out("Xb.Util.Barcode.GetString: " & ex.Message)
            Throw
        End Try
        IO.File.Delete(tmpName)

        Return result

    End Function

End Class
