Option Strict On

''' <summary>
''' �o�[�R�[�h�֘A�֐��Q
''' </summary>
''' <remarks>
''' �����Ӂ� ���{�ꕶ����ɂ͑Ή��o���Ȃ��B
''' </remarks>
Public Class Barcode


    ''' <summary>
    ''' �o�[�R�[�h�敪
    ''' </summary>
    ''' <remarks>
    ''' �o�[�R�[�h�̎�ނƓ���
    ''' http://www.media9.co.jp/m_tuhan/encyclopedia.html
    ''' 
    ''' �Ȃ��ADotNetBarcode�őΉ����Ă��邪 Jan8, Jan13, Code39, Qr ��4��̂݁B
    ''' </remarks>
    Public Enum Type

        ''' <summary>
        ''' JAN�R�[�h-13��
        ''' </summary>
        ''' <remarks>
        ''' ������F���p���l
        ''' �����@�F13���Œ�
        ''' ���ʃR�[�h�Ƃ��āAJIS�ɂ��K�i������Ă���B
        ''' ���[���b�p��EAN�A�A�����J��UPC�ƌ݊���������B
        ''' ���E���ʃR�[�h�����p�i�̂قڑS�ĂɎg�p����Ă���B
        ''' �H�i�A�G�݁A���ЁA�G���A�Ɠd�ȂǁB
        ''' </remarks>
        Jan13 = DotNetBarcode.Types.Jan13

        ''' <summary>
        ''' JAN�R�[�h-8��
        ''' </summary>
        ''' <remarks>
        ''' ������F���p���l
        ''' �����@�F8���Œ�
        ''' ���ʃR�[�h�Ƃ��āAJIS�ɂ��K�i������Ă���B
        ''' ���[���b�p��EAN�A�A�����J��UPC�ƌ݊���������B
        ''' ���E���ʃR�[�h�����p�i�̂قڑS�ĂɎg�p����Ă���B
        ''' �H�i�A�G�݁A���ЁA�G���A�Ɠd�ȂǁB
        ''' </remarks>
        Jan8 = DotNetBarcode.Types.Jan8

        ''' <summary>
        ''' CODE39
        ''' </summary>
        ''' <remarks>
        ''' ������F���p���l,���p�p��,�L��["-", " "(�X�y�[�X), "$", "/", "+", "%", "."],�X�^�[�g/�X�g�b�v�R�[�h["*"]�B
        ''' �����@�F��
        ''' �A���t�@�x�b�g��L���������邽�߁A�i�ԂȂǂ�\���ł���B
        ''' �H�Ɨp�o�[�R�[�h�Ƃ��đ����g���Ă���B���{�d�q�@�B�H�Ɖ�(EIAJ�j
        ''' </remarks>
        Code39 = DotNetBarcode.Types.Code39

        '''' <summary>
        '''' CODE128
        '''' </summary>
        '''' <remarks>
        '''' ������F���p���l,���p�p��(ASCII�R�[�h�͑S��OK)
        '''' �����@�F��
        '''' �����̕�����������B�����݂̂ŕ\���Ȃ�A�����Ƃ��T�C�Y���������Ȃ�B�i������12���ȏ�j
        '''' �č�UCC�̍��ڗp�����V���{��
        '''' </remarks>
        'Code128

        '''' <summary>
        '''' NW-7
        '''' </summary>
        '''' <remarks>
        '''' ������F���p���l,�L��["-", "$", ":", "/", "+", "."],�X�^�[�g/�X�g�b�v�R�[�h["a", "b", "c", "d", "d" ]
        '''' �����@�F��
        '''' �v�����^�ň�����₷�����W���[���\��
        '''' ��z�ւ̓`�[�ADPE�A�}���ق݂̑��o���Ǘ��ȂǁB
        '''' </remarks>
        'Nw7

        '''' <summary>
        '''' ITF-14��
        '''' </summary>
        '''' <remarks>
        '''' ������F���p���l
        '''' �����@�F14���Œ�
        '''' �����x�Ȉ󎚂��\�E�������i�R�[�h�Ƃ���1987�N��JIS�K�i������Ă���B
        '''' �W�������R�[�h�Ƃ��Ăi�h�r�����g���Ă���B
        '''' </remarks>
        'Itf14

        '''' <summary>
        '''' ITF-16��
        '''' </summary>
        '''' <remarks>
        '''' ������F���p���l
        '''' �����@�F16���Œ�
        '''' �����x�Ȉ󎚂��\�E�������i�R�[�h�Ƃ���1987�N��JIS�K�i������Ă���B
        '''' �W�������R�[�h�Ƃ��Ăi�h�r�����g���Ă���B
        '''' </remarks>
        'Itf16

        ''' <summary>
        ''' QR�R�[�h
        ''' </summary>
        ''' <remarks>
        ''' ������F�e��
        ''' �����@�F�ρA�p�� 1,523�����܂ŁA���� 644�����܂ŁB
        ''' ��\�I�ȂQ�����o�[�R�[�h�B���{�f���\�[�ŊJ�����ꂽ�R�[�h��2000�N6��15����ISO�̍��ۋK�i�Ƃ��Đ��肳�ꂽ
        ''' ���|�I�ȏ��ʂ������B
        ''' ���{�����ԍH�Ɖ�W��2�����R�[�h�Ƃ��č̗p����ȂǁA�e�핪��Ŋ��p���i��ł���B
        ''' </remarks>
        Qr = DotNetBarcode.Types.QRCode

        '''' <summary>
        '''' CODE93
        '''' </summary>
        '''' <remarks>
        '''' ������F���p���l,�L��["-", " "(�X�y�[�X), "$", "/", "+", "%", "." ],���p�p���啶��, �X�^�[�g/�X�g�b�v�R�[�h["��"], ���䕶��4��
        '''' �����@�F�H
        '''' ����L�����N�^��p���邱�Ƃɂ��t���A�X�L�[�̕\�����\�`�F�b�N�f�W�b�g��2�����Ă���
        '''' CODE39�̃o�[�R�[�h����Z������ׂɊJ�����ꂽ�o�[�R�[�h�ŁA�S�Ă�ASCII�R�[�h��\�����Ƃ��o����B
        '''' </remarks>
        'Code93

        '''' <summary>
        '''' INDUSTRIAL2 OF 5
        '''' </summary>
        '''' <remarks>
        '''' ������F���p���l
        '''' �����@�F��
        '''' �����Ǘ��ȂǂɎg�p����Ă���B
        '''' �ᖧ�x�Ńo�[�\�����P���ł���A�i�{�[������Ɍ����Ă���B��Ǘ��������������B
        '''' ���X�H�Ɨp�ɑ��p���ꂽ���A���������\�����ᖧ�x�̂��� CODE-39 �Ȃǂɒu��������ꂽ�B
        '''' </remarks>
        'Industrial2Of5

    End Enum


    ''' <summary>
    ''' �n���l������Ńo�[�R�[�h�摜�𐶐�����B
    ''' </summary>
    ''' <param name="value"></param>
    ''' <param name="type"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' DotNetBarcode�Ńo�[�R�[�h�摜�𐶐�����B
    ''' </remarks>
    Public Shared Function GetImage(ByVal value As String, _
                                    ByVal type As Xb.Media.Barcode.Type, _
                                    ByVal width As Integer, _
                                    ByVal height As Integer) As System.Drawing.Image

        Dim graphics As System.Drawing.Graphics, _
            image As Drawing.Bitmap = New Drawing.Bitmap(width, height), _
            builder As DotNetBarcode = New DotNetBarcode()

        'PictureBox����Graphics�I�u�W�F�N�g���擾�B
        graphics = Drawing.Graphics.FromImage(image)

        'Graphics�I�u�W�F�N�g�փo�[�R�[�h��`��B
        builder.Type = CType(type, DotNetBarcode.Types)
        builder.WriteBar(value, 0, 0, width, height, graphics)

        '�e�I�u�W�F�N�g��j���B
        graphics.Dispose()
        builder = Nothing

        Return image

    End Function


    ''' <summary>
    ''' �摜�t�@�C������o�[�R�[�h�f�[�^��������擾����B
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
            Xb.Util.Out("Xb.Util.Barcode.GetString: �摜�t�@�C����������܂���B")
            Throw New ArgumentException("�摜�t�@�C����������܂���B")
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
            Xb.Util.Out("Xb.Util.Barcode.GetString: �o�[�R�[�h��͂Ɏ��s���܂����B")
            Throw New ArgumentException("�o�[�R�[�h��͂Ɏ��s���܂����B")
        End Try

        If (nodes.Count <= 0) Then
            Xb.Util.Out("Xb.Util.Barcode.GetString: �o�[�R�[�h�����o�o���܂���ł����B")
            Throw New ArgumentException("�o�[�R�[�h�����o�o���܂���ł����B")
        End If

        Return nodes.Item(0).InnerText

    End Function


    ''' <summary>
    ''' �摜�I�u�W�F�N�g����o�[�R�[�h�f�[�^��������擾����B
    ''' </summary>
    ''' <param name="image"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' �摜�I�u�W�F�N�g����U�t�@�C���ɏ����o�������ƁAZBar�ɓǂ܂���B
    ''' </remarks>
    Public Shared Function GetString(ByRef image As System.Drawing.Image) As String

        If (image Is Nothing) Then
            Xb.Util.Out("Xb.Util.Barcode.GetString: �摜�I�u�W�F�N�g���s���ł��B")
            Throw New ArgumentException("�摜�I�u�W�F�N�g���s���ł��B")
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
