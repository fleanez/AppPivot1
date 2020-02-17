Imports System.IO
Imports EEPHASE

Module Utils

    'Private strFilePath As String
    Private objWriter As IO.StreamWriter

    ''' <summary>
    ''' Creates a new OpenPivot csv file in the specified path
    ''' </summary>
    ''' <param name="strCSVPath">Complete path including file name and extension</param>
    Public Sub CreateCSVfile(ByVal strCSVPath As String)
        Dim strHeader(7) As String
        strHeader(0) = "Model"
        strHeader(1) = "Company"
        strHeader(2) = "Parent"
        strHeader(3) = "Child"
        strHeader(4) = "Property"
        strHeader(5) = "Period"
        strHeader(6) = "Sample"
        strHeader(7) = "Value"
        'strHeader(7) = "Max Spare Capacity"
        'strHeader(8) = "Max Ramp Capacity"
        'strHeader(9) = "Max Response"
        'strHeader(10) = "Max Provision"

        File.Create(strCSVPath).Dispose()

        OpenFile(strCSVPath)
        AppendLine(strHeader)
        CloseFile()
    End Sub

    Private Sub AppendLine(ByVal strTextLine() As String)

        Dim strLine As String

        'If IO.File.Exists(strCSVPath) Then

        For Each strField As String In strTextLine

            strLine = ""

            'If value contains comma in the value then you have to perform this opertions
            Dim appendd = If(strField.Contains(","), String.Format("""{0}""", strField), strField)
            strLine = String.Format("{0}{1},", strLine, appendd)
            objWriter.Write(strLine)

        Next

        objWriter.Write(Environment.NewLine)

    End Sub

    ''' <summary>
    ''' Open up a stream to the given csv file
    ''' </summary>
    ''' <param name="strCSVPath">Full path to file</param>
    Public Sub OpenFile(strCSVPath As String)

        If (objWriter IsNot Nothing) Then
            objWriter.Close()
        End If
        objWriter = IO.File.AppendText(strCSVPath)

    End Sub

    ''' <summary>
    ''' Debug method for append data to opened file (first you must call openFile())
    ''' </summary>
    ''' <param name="strModel"></param>
    ''' <param name="strCompany"></param>
    ''' <param name="strParent"></param>
    ''' <param name="strChild"></param>
    ''' <param name="nPeriod"></param>
    ''' <param name="nSample"></param>
    ''' <param name="dMaxSpare"></param>
    ''' <param name="dMaxRamp"></param>
    ''' <param name="dMaxResponse"></param>
    ''' <param name="dMaxProvision"></param>
    Public Sub AppendDataToCsv(strModel As String, strCompany As String, strParent As String, strChild As String, strProperty As String, nPeriod As Integer, nSample As Integer, dMaxSpare As Double, dMaxRamp As Double, dMaxResponse As Double, dMaxProvision As Double)

        Dim strOut(10) As String
        strOut(0) = strModel
        strOut(1) = strCompany
        strOut(2) = strParent
        strOut(3) = strChild
        strOut(4) = strProperty
        strOut(5) = CStr(nPeriod)
        strOut(6) = CStr(nSample)
        strOut(7) = CStr(dMaxSpare)
        strOut(8) = CStr(dMaxRamp)
        strOut(9) = CStr(dMaxResponse)
        strOut(10) = CStr(dMaxProvision)
        AppendLine(strOut)

    End Sub

    ''' <summary>
    ''' Preferred method for append properties data to opened file (first you must call openFile())
    ''' </summary>
    ''' <param name="strModel"></param>
    ''' <param name="strCompany"></param>
    ''' <param name="strParent"></param>
    ''' <param name="strChild"></param>
    ''' <param name="strProperty"></param>
    ''' <param name="nPeriod"></param>
    ''' <param name="nSample"></param>
    ''' <param name="dValue"></param>
    Public Sub AppendDataToCsv(strModel As String, strCompany As String, strParent As String, strChild As String, strProperty As String, nPeriod As Integer, nSample As Integer, dValue As Double)
        Dim strOut(7) As String
        strOut(0) = strModel
        strOut(1) = strCompany
        strOut(2) = strParent
        strOut(3) = strChild
        strOut(4) = strProperty
        strOut(5) = CStr(nPeriod)
        strOut(6) = CStr(nSample)
        strOut(7) = CStr(dValue)
        AppendLine(strOut)
    End Sub

    ''' <summary>
    ''' Closes and flushes the connection to opened file (and sets to nothing the stream)
    ''' </summary>
    Public Sub CloseFile()
        If (objWriter IsNot Nothing) Then
            objWriter.Close()
            objWriter = Nothing
        End If
    End Sub

End Module
