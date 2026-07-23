use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: System.Lazy(Of T) Initialization & Thread Safety
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_lazy_initialization_deferred() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim initialized = False
        Dim lazyVal As New Lazy(Of Integer)(Function()
            initialized = True
            Return 42
        End Function)

        Console.WriteLine("IsCreatedBefore: " & lazyVal.IsValueCreated)
        Dim v = lazyVal.Value
        Console.WriteLine("IsCreatedAfter: " & lazyVal.IsValueCreated & "|Val=" & v)
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["IsCreatedBefore: False", "IsCreatedAfter: True|Val=42"]
    );
}

#[test]
fn test_vb_lazy_value_cached_on_subsequent_calls() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim counter = 0
        Dim lazyVal As New Lazy(Of Integer)(Function()
            counter += 1
            Return counter * 10
        End Function)

        Dim v1 = lazyVal.Value
        Dim v2 = lazyVal.Value
        Console.WriteLine(v1 & "|" & v2 & "|Counter=" & counter)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10|10|Counter=1"]);
}

#[test]
fn test_vb_lazy_thread_safety_mode_publication_only() {
    let src = r#"
Imports System
Imports System.Threading

Module Program
    Sub Main()
        Dim lazyVal As New Lazy(Of String)(Function() "ThreadSafeVal", LazyThreadSafetyMode.PublicationOnly)
        Console.WriteLine(lazyVal.Value)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["ThreadSafeVal"]);
}

#[test]
fn test_vb_lazy_thread_safety_mode_execution_and_publication() {
    let src = r#"
Imports System
Imports System.Threading

Module Program
    Sub Main()
        Dim lazyVal As New Lazy(Of Integer)(Function() 100, LazyThreadSafetyMode.ExecutionAndPublication)
        Console.WriteLine(lazyVal.IsValueCreated & "|" & lazyVal.Value)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["False|100"]);
}

#[test]
fn test_vb_lazy_thread_safety_mode_none() {
    let src = r#"
Imports System
Imports System.Threading

Module Program
    Sub Main()
        Dim lazyVal As New Lazy(Of String)(Function() "SingleThreaded", LazyThreadSafetyMode.None)
        Console.WriteLine(lazyVal.Value)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["SingleThreaded"]);
}

#[test]
fn test_vb_lazy_default_constructor_calls_parameterless_ctor() {
    let src = r#"
Imports System

Class Widget
    Public Property Name As String = "DefaultWidget"
End Class

Module Program
    Sub Main()
        Dim lazyVal As New Lazy(Of Widget)()
        Console.WriteLine(lazyVal.IsValueCreated & "|" & lazyVal.Value.Name)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["False|DefaultWidget"]);
}

#[test]
fn test_vb_lazy_exception_cached_on_failure() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim attempts = 0
        Dim lazyVal As New Lazy(Of String)(Function()
            attempts += 1
            Throw New InvalidOperationException("Fail " & attempts)
        End Function)

        Try
            Dim v = lazyVal.Value
        Catch ex1 As InvalidOperationException
            Console.WriteLine(ex1.Message)
        End Try

        Try
            Dim v2 = lazyVal.Value
        Catch ex2 As InvalidOperationException
            Console.WriteLine(ex2.Message & "|Attempts=" & attempts)
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Fail 1", "Fail 1|Attempts=1"]);
}

#[test]
fn test_vb_lazy_publication_only_does_not_cache_exception() {
    let src = r#"
Imports System
Imports System.Threading

Module Program
    Sub Main()
        Dim attempts = 0
        Dim lazyVal As New Lazy(Of String)(Function()
            attempts += 1
            If attempts = 1 Then Throw New InvalidOperationException("Fail 1")
            Return "Success"
        End Function, LazyThreadSafetyMode.PublicationOnly)

        Try
            Dim v = lazyVal.Value
        Catch ex As InvalidOperationException
            Console.WriteLine("First Attempt Failed")
        End Try

        Dim vSuccess = lazyVal.Value
        Console.WriteLine(vSuccess & "|Attempts=" & attempts)
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["First Attempt Failed", "Success|Attempts=2"]
    );
}

#[test]
fn test_vb_lazy_boolean_constructor_is_thread_safe() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        ' Lazy(Of T)(isThreadSafe:=True)
        Dim lazyVal As New Lazy(Of Integer)(Function() 999, True)
        Console.WriteLine(lazyVal.Value)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["999"]);
}

#[test]
fn test_vb_lazy_to_string_representation() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim lazyVal As New Lazy(Of Integer)(Function() 777)
        Dim strBefore = lazyVal.ToString()
        Dim val = lazyVal.Value
        Dim strAfter = lazyVal.ToString()
        Console.WriteLine(strBefore & "|" & strAfter)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Value is not created.|777"]);
}

#[test]
fn test_vb_lazy_custom_reference_type_factory() {
    let src = r#"
Imports System

Class Config
    Public Property Port As Integer
End Class

Module Program
    Sub Main()
        Dim lazyConfig As New Lazy(Of Config)(Function() New Config With {.Port = 8080})
        Console.WriteLine(lazyConfig.Value.Port)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["8080"]);
}

#[test]
fn test_vb_lazy_value_type_struct_factory() {
    let src = r#"
Imports System

Structure Point2D
    Public X As Integer
    Public Y As Integer
End Structure

Module Program
    Sub Main()
        Dim lazyPoint As New Lazy(Of Point2D)(Function() New Point2D With {.X = 10, .Y = 20})
        Console.WriteLine(lazyPoint.Value.X & "," & lazyPoint.Value.Y)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10,20"]);
}

#[test]
fn test_vb_lazy_null_return_value_valid() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim lazyVal As New Lazy(Of String)(Function() CType(Nothing, String))
        Dim str = lazyVal.Value
        Console.WriteLine(lazyVal.IsValueCreated & "|" & (str Is Nothing))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True"]);
}

#[test]
fn test_vb_lazy_recursion_exception_detection() {
    let src = r#"
Imports System

Module Program
    Private recLazy As Lazy(Of Integer)

    Sub Main()
        recLazy = New Lazy(Of Integer)(Function() recLazy.Value + 1)
        Try
            Dim val = recLazy.Value
        Catch ex As InvalidOperationException
            Console.WriteLine("Recursive Lazy Initialization Caught")
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Recursive Lazy Initialization Caught"]);
}

#[test]
fn test_vb_lazy_list_collection_initialization() {
    let src = r#"
Imports System
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim lazyList As New Lazy(Of List(Of String))(Function() New List(Of String) From {"A", "B", "C"})
        Console.WriteLine(lazyList.Value.Count & ":" & String.Join(",", lazyList.Value))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3:A,B,C"]);
}

#[test]
fn test_vb_lazy_null_value_factory_delegate_throws() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Try
            Dim lazyVal As New Lazy(Of String)(CType(Nothing, Func(Of String)))
        Catch ex As ArgumentNullException
            Console.WriteLine("ArgumentNullException Caught on Null Delegate")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["ArgumentNullException Caught on Null Delegate"]
    );
}

#[test]
fn test_vb_lazy_thread_local_storage_simulation() {
    let src = r#"
Imports System
Imports System.Threading

Module Program
    Sub Main()
        Dim threadLocalVal As New ThreadLocal(Of Integer)(Function() Thread.CurrentThread.ManagedThreadId * 10)
        Console.WriteLine(threadLocalVal.IsValueCreated & "|" & (threadLocalVal.Value > 0))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["False|True"]);
}

#[test]
fn test_vb_lazy_linq_deferred_evaluation() {
    let src = r#"
Imports System
Imports System.Linq

Module Program
    Sub Main()
        Dim lazyNumbers As New Lazy(Of IEnumerable(Of Integer))(Function() Enumerable.Range(1, 5))
        Dim sum = lazyNumbers.Value.Sum()
        Console.WriteLine(sum)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["15"]);
}

#[test]
fn test_vb_lazy_nested_lazy_structures() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim outer As New Lazy(Of Lazy(Of String))(Function() New Lazy(Of String)(Function() "NestedLazyVal"))
        Console.WriteLine(outer.Value.Value)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["NestedLazyVal"]);
}

#[test]
fn test_vb_lazy_disposal_pattern_on_value() {
    let src = r#"
Imports System
Imports System.IO

Module Program
    Sub Main()
        Dim lazyStream As New Lazy(Of MemoryStream)(Function() New MemoryStream())
        lazyStream.Value.WriteByte(123)
        Console.WriteLine(lazyStream.Value.Length)
        lazyStream.Value.Dispose()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1"]);
}
