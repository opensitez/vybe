' vybe-test: vb/vb_objects_collections/d34_stack_process_until_empty
' origin: languages/vb/tests/vb/vb_objects_collections_test.rs

Dim s As New Stack(Of Integer)
s.Push(10)
s.Push(20)
s.Push(30)
Dim total As Integer = 0
Do While s.Count > 0
    total = total + s.Pop()
Loop
Console.WriteLine(total)
