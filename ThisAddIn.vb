Imports rssfeedPower.feedPowerPoint
Imports PowerPoint = Microsoft.Office.Interop.PowerPoint
Imports Office = Microsoft.Office.Core
Imports System.Net
Imports Microsoft.Office.Interop.PowerPoint
Imports System.Timers
Imports System.Threading
Public Class ThisAddIn
    Private pres As PowerPoint.Presentation
    Private Sub ThisAddIn_Startup() Handles Me.Startup
        If Not System.IO.Directory.Exists("C:\fetched") Then
            System.IO.Directory.CreateDirectory("C:\fetched")
        End If
        Dim Start As DateTime = DateTime.Now.AddSeconds(0)
        Dim Interval As TimeSpan = TimeSpan.FromSeconds(60)
        StartTimer(Start, Interval, AddressOf TriggerNewPresentation) ' Trigger a new presentation every 60 seconds
    End Sub
    ' Trigger a new prsentation  when called from main thread
#Region "For New Presenation"
    Private Function TriggerNewPresentation() As Boolean
        Dim result As Boolean = False
        Try
            Me.Application.ActivePresentation.Close()
            pres = Me.Application.Presentations.Add(Office.MsoTriState.msoTrue)
            Dim Start As DateTime = DateTime.Now.AddSeconds(0)
            Dim Interval As TimeSpan = TimeSpan.FromSeconds(10)
            StartTimer(Start, Interval, AddressOf addImages)
            Dim images As String() = System.IO.Directory.GetFiles("C:\fetched")
            If images.Length > 0 Then
                Dim sec As String = System.DateTime.Now.Second.ToString()
                Dim min As String = System.DateTime.Now.Minute.ToString()
                Dim hr As String = System.DateTime.Now.Hour.ToString()
                Dim [date] As String = System.DateTime.Now.Day.ToString()
                Dim mon As String = System.DateTime.Now.Month.ToString()
                Dim yr As String = System.DateTime.Now.Year.ToString()
                Dim name As String = yr + "-" + mon + "-" + [date] + "-" + hr + "-" + min + "-" + sec
                Me.Application.ActivePresentation.SaveAs("C:\Users\Public\Documents\sandeep_" + name, PowerPoint.PpSaveAsFileType.ppSaveAsDefault, Office.MsoTriState.msoTriStateMixed)
                'For Each img As String In images
                '    System.IO.File.Delete(img)
                'Next
            End If
            result = True
        Catch
        End Try
        Return result
    End Function
   
#End Region

#Region "Timer Thread"
    Private Delegate Sub voidFunc(Of P1, P2, P3)(p1 As P1, p2 As P2, p3 As P3)
    Public Sub StartTimer(startTime As DateTime, interval As TimeSpan, action As Func(Of Boolean))
        Dim Timer As voidFunc(Of DateTime, TimeSpan, Func(Of Boolean)) = AddressOf TimedThread
        Timer.BeginInvoke(startTime, interval, action, Nothing, Nothing)
    End Sub
    Private Sub TimedThread(startTime As DateTime, interval As TimeSpan, action As Func(Of Boolean))
        Dim keepRunning As Boolean = True
        Dim NextExecute As DateTime = startTime
        While keepRunning
            If DateTime.Now > NextExecute Then
                keepRunning = action.Invoke()
                NextExecute = NextExecute.Add(interval)
            End If
            Thread.Sleep(1000)
        End While
    End Sub
#End Region

    ' Fetch Images from Server....Sandeep
#Region "Fetching Images from Rss Feed"
    Private Function addImages() As Boolean
        Dim master As PowerPoint.Master
        Dim slide As PowerPoint.Slide
        Dim cl As PowerPoint.CustomLayout
        Dim result As Boolean = False
        Dim fetchingThread As New Thread(New ThreadStart(AddressOf fetchImages))
        fetchingThread.Start()
        fetchingThread.Join()
        Dim images As String() = System.IO.Directory.GetFiles("C:\fetched")
        master = pres.SlideMaster
        cl = master.CustomLayouts(PpSlideLayout.ppLayoutChartAndText)
        If images.Length > 0 Then
            Dim i As Integer = 0
            While i <= images.Length - 1
                slide = pres.Slides.AddSlide(1, cl)
                slide.Shapes.AddPicture(images(i), Office.MsoTriState.msoTrue, Office.MsoTriState.msoTrue, 10, 10, 700, _
                 530)
                System.Math.Max(System.Threading.Interlocked.Increment(i), i - 1)
            End While
        End If
        result = True
        Return result
    End Function
    ' Fetch Images from Server....Sandeep
    Private Sub fetchImages()
        Dim result As Boolean = False
        Dim item As New RssFeedItem()
        Dim list As New List(Of RssFeedItem)()
        list = RssManager.ReadFeed("http://83.218.82.67:81/rss.aspx?cid=124")
        Try
            For Each feeditem As RssFeedItem In list
                Dim imgUri As New Uri(feeditem.imglnk.ToString())
                Dim webClient As New WebClient()
                Dim lnkArray As String() = feeditem.imglnk.ToString().Split("/")
                Dim iImage As String = lnkArray(5)
                Dim localFileName As String = "C:\fetched\" + iImage
                If Not System.IO.File.Exists("C:\fetched\" + iImage) Then
                    Try
                        webClient.DownloadFile(imgUri, localFileName)
                    Catch
                    End Try
                End If
            Next
            result = True
        Catch
        End Try
    End Sub
#End Region
    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub

End Class
