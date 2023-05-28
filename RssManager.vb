Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Net
Imports System.Data
Namespace feedPowerPoint
    ''' <summary>
    ''' RSS manager to read rss feeds
    ''' </summary>
    Public Class RssManager
        ''' <summary>
        ''' Reads the relevant Rss feed and returns a list off RssFeedItems
        ''' </summary>
        ''' <param name="url"></param>
        ''' <returns></returns>
        Public Shared Function ReadFeed(url As String) As List(Of RssFeedItem)
            'create a new list of the rss feed items to return
            Dim rssFeedItems As New List(Of RssFeedItem)()
            'create an http request which will be used to retrieve the rss feed
            Dim rssFeed As HttpWebRequest = DirectCast(WebRequest.Create(url), HttpWebRequest)
            'use a dataset to retrieve the rss feed
            Using rssData As New DataSet()
                'read the xml from the stream of the web request
                rssData.ReadXml(rssFeed.GetResponse().GetResponseStream())
                Dim tbls As DataTableCollection
                tbls = rssData.Tables
                'loop through the rss items in the dataset and populate the list of rss feed items
                For Each dataRow As DataRow In rssData.Tables("guid").Rows
                    'ChannelId = Convert.ToInt32(dataRow["channel_Id"]),
                    'Description = Convert.ToString(dataRow["author"]),
                    'ItemId = Convert.ToInt32(dataRow["item_Id"]),
                    'Link = Convert.ToString(dataRow["link"]),
                    'Enclosure = Convert.ToString(dataRow["guid"]),
                    'PublishDate = Convert.ToDateTime(dataRow["pubDate"]),
                    'Title = Convert.ToString(dataRow["title"])
                    rssFeedItems.Add(New RssFeedItem())
                Next
            End Using
            'return the rss feed items
            Return rssFeedItems
        End Function
    End Class
End Namespace