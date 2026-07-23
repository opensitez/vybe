use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Async/Await & LINQ Reactive Query Pipelines
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_async_linq_select_async_data_transformation() {
    let src = r#"
Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Threading.Tasks

Module Program
    Private Async Function FetchDataAsync(id As Integer) As Task(Of String)
        Await Task.Yield()
        Return "Item_" & id
    End Function

    Sub Main()
        Dim ids As Integer() = {1, 2, 3}
        Dim tasks = ids.Select(Function(i) FetchDataAsync(i)).ToArray()
        Task.WaitAll(tasks)

        Dim results = tasks.Select(Function(t) t.Result).ToList()
        Console.WriteLine(String.Join(",", results))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Item_1,Item_2,Item_3"]);
}

#[test]
fn test_vb_async_linq_where_filter_pipeline() {
    let src = r#"
Imports System.Collections.Generic
Imports System.Linq
Imports System.Threading.Tasks

Module Program
    Private Async Function IsValidAsync(num As Integer) As Task(Of Boolean)
        Await Task.Yield()
        Return num Mod 2 = 0
    End Function

    Sub Main()
        Dim numbers As Integer() = {10, 15, 20, 25, 30}
        ' Filter asynchronously
        Dim tasks = numbers.Select(Async Function(n) New With {.Val = n, .Keep = Await IsValidAsync(n)}).ToArray()
        Task.WaitAll(tasks)

        Dim evens = tasks.Where(Function(x) x.Result.Keep).Select(Function(x) x.Result.Val)
        Console.WriteLine(String.Join(",", evens))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10,20,30"]);
}

#[test]
fn test_vb_async_linq_aggregate_sum_pipeline() {
    let src = r#"
Imports System.Linq
Imports System.Threading.Tasks

Module Program
    Private Async Function ComputeValueAsync(x As Integer) As Task(Of Integer)
        Await Task.Yield()
        Return x * 10
    End Function

    Sub Main()
        Dim inputs As Integer() = {1, 2, 3, 4}
        Dim tasks = inputs.Select(Function(x) ComputeValueAsync(x)).ToArray()
        Task.WaitAll(tasks)

        Dim totalSum = tasks.Sum(Function(t) t.Result)
        Console.WriteLine(totalSum)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["100"]);
}

#[test]
fn test_vb_async_linq_group_by_pipeline() {
    let src = r#"
Imports System.Collections.Generic
Imports System.Linq
Imports System.Threading.Tasks

Module Program
    Class Metric
        Public Category As String
        Public Value As Integer
    End Class

    Private Async Function GetMetricsAsync() As Task(Of List(Of Metric))
        Await Task.Yield()
        Return New List(Of Metric) From {
            New Metric With {.Category = "CPU", .Value = 50},
            New Metric With {.Category = "RAM", .Value = 70},
            New Metric With {.Category = "CPU", .Value = 60}
        }
    End Function

    Sub Main()
        Dim t = GetMetricsAsync()
        t.Wait()

        Dim grouped = t.Result.GroupBy(Function(m) m.Category)
        For Each g In grouped.OrderBy(Function(g) g.Key)
            Console.WriteLine(g.Key & ":" & g.Average(Function(m) m.Value))
        Next
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["CPU:55", "RAM:70"]);
}

#[test]
fn test_vb_async_linq_first_or_default_async_predicate() {
    let src = r#"
Imports System.Linq
Imports System.Threading.Tasks

Module Program
    Private Async Function CheckMatchAsync(s As String) As Task(Of Boolean)
        Await Task.Yield()
        Return s.StartsWith("B")
    End Function

    Sub Main()
        Dim items As String() = {"Apple", "Banana", "Cherry"}
        Dim tasks = items.Select(Async Function(item) New With {.Text = item, .IsMatch = Await CheckMatchAsync(item)}).ToArray()
        Task.WaitAll(tasks)

        Dim firstMatch = tasks.FirstOrDefault(Function(t) t.Result.IsMatch)
        Console.WriteLine(If(firstMatch IsNot Nothing, firstMatch.Result.Text, "None"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Banana"]);
}

#[test]
fn test_vb_async_linq_task_when_all_projection() {
    let src = r#"
Imports System.Linq
Imports System.Threading.Tasks

Module Program
    Private Async Function SquareAsync(n As Integer) As Task(Of Integer)
        Await Task.Yield()
        Return n * n
    End Function

    Sub Main()
        Dim numbers As Integer() = {2, 3, 4}
        Dim whenAllTask = Task.WhenAll(numbers.Select(Function(n) SquareAsync(n)))
        whenAllTask.Wait()

        Console.WriteLine(String.Join("+", whenAllTask.Result))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["4+9+16"]);
}

#[test]
fn test_vb_async_linq_flat_map_select_many() {
    let src = r#"
Imports System.Collections.Generic
Imports System.Linq
Imports System.Threading.Tasks

Module Program
    Private Async Function GetSubItemsAsync(category As String) As Task(Of String())
        Await Task.Yield()
        Return New String() {category & "-1", category & "-2"}
    End Function

    Sub Main()
        Dim categories As String() = {"CatA", "CatB"}
        Dim tasks = categories.Select(Function(c) GetSubItemsAsync(c)).ToArray()
        Task.WaitAll(tasks)

        Dim flattened = tasks.SelectMany(Function(t) t.Result).ToList()
        Console.WriteLine(String.Join(",", flattened))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["CatA-1,CatA-2,CatB-1,CatB-2"]);
}

#[test]
fn test_vb_async_linq_zip_join_projections() {
    let src = r#"
Imports System.Linq
Imports System.Threading.Tasks

Module Program
    Private Async Function FetchNamesAsync() As Task(Of String())
        Await Task.Yield()
        Return New String() {"Alice", "Bob"}
    End Function

    Private Async Function FetchScoresAsync() As Task(Of Integer())
        Await Task.Yield()
        Return New Integer() {90, 85}
    End Function

    Sub Main()
        Dim tNames = FetchNamesAsync()
        Dim tScores = FetchScoresAsync()
        Task.WaitAll(tNames, tScores)

        Dim zipped = tNames.Result.Zip(tScores.Result, Function(name, score) name & "=" & score)
        Console.WriteLine(String.Join("|", zipped))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Alice=90|Bob=85"]);
}

#[test]
fn test_vb_async_linq_order_by_async_computed_key() {
    let src = r#"
Imports System.Linq
Imports System.Threading.Tasks

Module Program
    Private Async Function ComputeWeightAsync(s As String) As Task(Of Integer)
        Await Task.Yield()
        Return s.Length
    End Function

    Sub Main()
        Dim words As String() = {"Elephant", "Cat", "Giraffe"}
        Dim tasks = words.Select(Async Function(w) New With {.Word = w, .Weight = Await ComputeWeightAsync(w)}).ToArray()
        Task.WaitAll(tasks)

        Dim sortedWords = tasks.Select(Function(t) t.Result).OrderBy(Function(x) x.Weight).Select(Function(x) x.Word)
        Console.WriteLine(String.Join(",", sortedWords))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Cat,Giraffe,Elephant"]);
}

#[test]
fn test_vb_async_linq_chunking_and_batching() {
    let src = r#"
Imports System.Linq
Imports System.Threading.Tasks

Module Program
    Private Async Function ProcessBatchAsync(batch As Integer()) As Task(Of Integer)
        Await Task.Yield()
        Return batch.Sum()
    End Function

    Sub Main()
        Dim items As Integer() = {1, 2, 3, 4, 5, 6}
        Dim chunks = items.Chunk(2)
        Dim tasks = chunks.Select(Function(c) ProcessBatchAsync(c)).ToArray()
        Task.WaitAll(tasks)

        Console.WriteLine(String.Join(",", tasks.Select(Function(t) t.Result)))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3,7,11"]);
}

#[test]
fn test_vb_async_linq_retry_policy_pipeline() {
    let src = r#"
Imports System
Imports System.Threading.Tasks

Module Program
    Private attempts As Integer = 0

    Private Async Function UnreliableAsync() As Task(Of String)
        Await Task.Yield()
        attempts += 1
        If attempts < 3 Then Throw New InvalidOperationException("Fail " & attempts)
        Return "Success"
    End Function

    Sub Main()
        Dim t = Task.Run(Async Function()
            For i As Integer = 1 To 5
                Try
                    Return Await UnreliableAsync()
                Catch ex As Exception
                End Try
            Next
            Return "FailedAll"
        End Function)
        t.Wait()
        Console.WriteLine(t.Result & "|Attempts=" & attempts)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Success|Attempts=3"]);
}

#[test]
fn test_vb_async_linq_cancellation_token_propagation() {
    let src = r#"
Imports System
Imports System.Linq
Imports System.Threading
Imports System.Threading.Tasks

Module Program
    Private Async Function LongWorkAsync(n As Integer, ct As CancellationToken) As Task(Of Integer)
        ct.ThrowIfCancellationRequested()
        Await Task.Yield()
        Return n * 2
    End Function

    Sub Main()
        Dim cts As New CancellationTokenSource()
        cts.Cancel()

        Dim tasks = Enumerable.Range(1, 5).Select(Function(n) LongWorkAsync(n, cts.Token)).ToArray()

        Try
            Task.WaitAll(tasks)
        Catch ex As AggregateException
            Console.WriteLine("AggregateException Caught on Cancelled Pipeline")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["AggregateException Caught on Cancelled Pipeline"]
    );
}

#[test]
fn test_vb_async_linq_take_while_condition() {
    let src = r#"
Imports System.Linq
Imports System.Threading.Tasks

Module Program
    Private Async Function IsUnderLimitAsync(val As Integer) As Task(Of Boolean)
        Await Task.Yield()
        Return val < 100
    End Function

    Sub Main()
        Dim items As Integer() = {10, 50, 120, 30}
        Dim tasks = items.Select(Async Function(x) New With {.Val = x, .Valid = Await IsUnderLimitAsync(x)}).ToArray()
        Task.WaitAll(tasks)

        Dim validSequence = tasks.Select(Function(t) t.Result).TakeWhile(Function(x) x.Valid).Select(Function(x) x.Val)
        Console.WriteLine(String.Join(",", validSequence))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10,50"]);
}

#[test]
fn test_vb_async_linq_dictionary_projection() {
    let src = r#"
Imports System.Linq
Imports System.Threading.Tasks

Module Program
    Private Async Function GetKeyValuePairAsync(id As Integer) As Task(Of KeyValuePair(Of Integer, String))
        Await Task.Yield()
        Return New KeyValuePair(Of Integer, String)(id, "Code_" & id)
    End Function

    Sub Main()
        Dim ids As Integer() = {101, 102}
        Dim tasks = ids.Select(Function(i) GetKeyValuePairAsync(i)).ToArray()
        Task.WaitAll(tasks)

        Dim dict = tasks.Select(Function(t) t.Result).ToDictionary(Function(kvp) kvp.Key, Function(kvp) kvp.Value)
        Console.WriteLine(dict(101) & "|" & dict(102))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Code_101|Code_102"]);
}

#[test]
fn test_vb_async_linq_distinct_async_results() {
    let src = r#"
Imports System.Linq
Imports System.Threading.Tasks

Module Program
    Private Async Function NormalizeAsync(s As String) As Task(Of String)
        Await Task.Yield()
        Return s.Trim().ToUpper()
    End Function

    Sub Main()
        Dim raw As String() = {"apple", " Apple ", "APPLE", "banana"}
        Dim tasks = raw.Select(Function(s) NormalizeAsync(s)).ToArray()
        Task.WaitAll(tasks)

        Dim unique = tasks.Select(Function(t) t.Result).Distinct().OrderBy(Function(x) x)
        Console.WriteLine(String.Join(",", unique))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["APPLE,BANANA"]);
}

#[test]
fn test_vb_async_linq_all_any_async_predicates() {
    let src = r#"
Imports System.Linq
Imports System.Threading.Tasks

Module Program
    Private Async Function IsPositiveAsync(n As Integer) As Task(Of Boolean)
        Await Task.Yield()
        Return n > 0
    End Function

    Sub Main()
        Dim numbers As Integer() = {1, 5, 10}
        Dim tasks = numbers.Select(Async Function(n) Await IsPositiveAsync(n)).ToArray()
        Task.WaitAll(tasks)

        Dim allPositive = tasks.All(Function(t) t.Result)
        Console.WriteLine(allPositive)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_async_linq_exception_handling_in_select_projection() {
    let src = r#"
Imports System
Imports System.Linq
Imports System.Threading.Tasks

Module Program
    Private Async Function SafeDivideAsync(n As Integer) As Task(Of Double)
        Await Task.Yield()
        If n = 0 Then Throw New DivideByZeroException("Cannot divide by zero")
        Return 100.0 / n
    End Function

    Sub Main()
        Dim numbers As Integer() = {10, 0, 5}
        Dim tasks = numbers.Select(Function(n) SafeDivideAsync(n)).ToArray()

        Dim results As New System.Collections.Generic.List(Of String)()
        For Each t In tasks
            Try
                t.Wait()
                results.Add(t.Result.ToString("F0"))
            Catch ex As AggregateException
                results.Add("Error")
            End Try
        Next
        Console.WriteLine(String.Join(",", results))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10,Error,20"]);
}

#[test]
fn test_vb_async_linq_task_when_any_first_responder() {
    let src = r#"
Imports System.Linq
Imports System.Threading.Tasks

Module Program
    Private Async Function SlowTaskAsync() As Task(Of String)
        Await Task.Delay(50)
        Return "Slow"
    End Function

    Private Async Function FastTaskAsync() As Task(Of String)
        Await Task.Yield()
        Return "Fast"
    End Function

    Sub Main()
        Dim tSlow = SlowTaskAsync()
        Dim tFast = FastTaskAsync()

        Dim winner = Task.WhenAny(tSlow, tFast)
        winner.Wait()
        Console.WriteLine(winner.Result.Result)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Fast"]);
}

#[test]
fn test_vb_async_linq_complex_nested_pipeline() {
    let src = r#"
Imports System.Linq
Imports System.Threading.Tasks

Module Program
    Private Async Function TransformRecordAsync(val As Integer) As Task(Of Integer)
        Await Task.Yield()
        Return val * 3
    End Function

    Sub Main()
        Dim input As Integer() = {1, 2, 3, 4, 5}
        Dim pipelineTask = Task.Run(Async Function()
            ' Filter even -> Multiply by 3 -> Sum
            Dim evens = input.Where(Function(n) n Mod 2 = 0)
            Dim tasks = evens.Select(Function(n) TransformRecordAsync(n)).ToArray()
            Await Task.WhenAll(tasks)
            Return tasks.Sum(Function(t) t.Result)
        End Function)

        pipelineTask.Wait()
        Console.WriteLine(pipelineTask.Result)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["18"]);
}

#[test]
fn test_vb_async_linq_empty_source_sequence() {
    let src = r#"
Imports System.Linq
Imports System.Threading.Tasks

Module Program
    Sub Main()
        Dim emptyItems As Integer() = {}
        Dim tasks = emptyItems.Select(Async Function(n) Await Task.FromResult(n * 2)).ToArray()
        Task.WaitAll(tasks)
        Console.WriteLine(tasks.Length)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0"]);
}
