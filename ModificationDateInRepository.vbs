Dim rc : Set rc = RepositoryConnection
rc.Open "Repository", "p-lucasmisilva", "@4ELM6qza", "p-lucasmisilva", "@4ELM6qza"
Output "Connected!"
ListChildren rc
'Create Folder
Dim TargetFolder : Set TargetFolder = rc.CreateFolder("Modelo de Dados")
ListChildren rc
rc.Close
Output "Disconnected!"
'List repository contents
Sub ListChildren(rc)
Output "Repository Contents:"
For each c in rc.ChildObjects
 Output "* " & c.Name & " (" & c.metaclass.PublicName & " - Modified: " & c.ModificationDateInRepository & ")"
Next
End Sub

