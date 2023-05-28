Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Namespace feedPowerPoint
    ''' <summary>
    ''' RSS feed item entity
    ''' </summary>
    Public Class RssFeedItem
        ''' <summary>
        ''' Gets or sets the title
        ''' </summary>
        Private _Title As String
        Public Property Title() As String
            Get
                Return _Title
            End Get
            Set(value As String)
            End Set
        End Property
        ''' <summary>
        ''' Gets or sets the description
        ''' </summary>
        Private _Description As String
        Public Property Description() As String
            Get
                Return _Description
            End Get
            Set(value As String)
            End Set
        End Property
        ''' <summary>
        ''' Gets or sets the link
        ''' </summary>
        Private _Link As String
        Public Property Link() As String
            Get
                Return _Link
            End Get
            Set(value As String)
            End Set
        End Property
        ''' <summary>
        ''' Gets or sets the link
        ''' </summary>
        Private _imglnk As String
        Public Property imglnk() As String
            Get
                Return _imglnk
            End Get
            Set(value As String)
            End Set
        End Property
        ''' <summary>
        ''' Gets or sets the item id
        ''' </summary>
        Private _ItemId As Integer
        Public Property ItemId() As Integer
            Get
                Return _ItemId
            End Get
            Set(value As Integer)
            End Set
        End Property
        ''' <summary>
        ''' Gets or sets the publish date
        ''' </summary>
        Private _PublishDate As DateTime
        Public Property PublishDate() As DateTime
            Get
                Return _PublishDate
            End Get
            Set(value As DateTime)
            End Set
        End Property
        ''' <summary>
        ''' Gets or sets the channel id
        ''' </summary>
        Private _ChannelID As Integer
        Public Property ChannelId() As Integer
            Get
                Return _ChannelID
            End Get
            Set(value As Integer)
            End Set
        End Property
    End Class
End Namespace