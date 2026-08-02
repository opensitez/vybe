' vybe-test: vb/vb_objects_collections/d31_queue_process_until_empty
' origin: languages/vb/tests/vb/vb_objects_collections_test.rs

Dim q As New Queue(Of Integer)
q.Enqueue(1)
q.Enqueue(2)
q.Enqueue(3)
Dim total As Integer = 0
Do While q.Count > 0
    total = total + q.Dequeue()
Loop
Console.WriteLine(total)
