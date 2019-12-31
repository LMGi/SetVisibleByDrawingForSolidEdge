Imports System.Runtime.InteropServices
Module MainModule
    Private mSolidApp As Object
    Private mDraftDoc As Object
    Private mAsmDoc As Object
    Sub Main()
        Try
            OleMessageFilter.Register()
            mSolidApp = Marshal.GetActiveObject("SolidEdge.Application")
            mSolidApp.DelayCompute = True
            mSolidApp.Interactive = False
            mDraftDoc = mSolidApp.ActiveDocument
            Dim modLink = mDraftDoc.ModelLinks.Item(1)
            mAsmDoc = mSolidApp.Documents.Open(modLink.Filename)
            MatchVisibility(modLink.ModelNodes, mAsmDoc.Occurrences)
        Catch ex As Exception
            MsgBox(ex.Message, "Fail")
        Finally
            mSolidApp.DelayCompute = False
            mSolidApp.Interactive = True
            OleMessageFilter.Unregister()
        End Try
    End Sub
    Private Sub MatchVisibility(modelNodes As Object, occs As Object)
        Dim o As Object = Nothing
        Dim n As Object = Nothing
        For Each n In modelNodes
            Dim componentName As String = n.ComponentName
            Dim fileName As String = n.FileName
            For Each o In occs
                'Find by name because you can reorder assemblies
                If String.Compare(o.Name, componentName, True) = 0 Then
                    o.Visible = n.Visible
                    Exit For
                End If
            Next
            If fileName.EndsWith(".asm") Then
                If n.Visible Then 'No need to traverse if the entire asm is hidden
                    MatchVisibility(n.ModelNodes, o.SubOccurrences)
                End If
            End If
        Next
    End Sub
End Module