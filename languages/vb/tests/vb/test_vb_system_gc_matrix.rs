use super::helpers::run_vb;

#[test]
fn gc_collect_and_generation_bounds() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim max As Integer = GC.MaxGeneration
        GC.Collect()
        GC.WaitForPendingFinalizers()
        GC.Collect()
        Console.WriteLine(max >= 0)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True"]);
}

#[test]
fn gc_get_total_memory_returns_non_negative() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim memory As Long = GC.GetTotalMemory(False)
        Console.WriteLine(memory >= 0)
        GC.GetTotalMemory(True)
        Console.WriteLine(memory >= 0)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn gc_collection_counts_are_queryable() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim before0 As Integer = GC.CollectionCount(0)
        Dim before1 As Integer = GC.CollectionCount(1)

        Dim arr(9_999_999) As Byte
        arr = Nothing
        GC.Collect()
        GC.WaitForPendingFinalizers()

        Dim after0 As Integer = GC.CollectionCount(0)
        Dim after1 As Integer = GC.CollectionCount(1)

        Console.WriteLine(after0 >= before0)
        Console.WriteLine(after1 >= before1)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn gc_get_generation_of_null_is_minus_one() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim value As Integer = GC.GetGeneration(Nothing)
        Console.WriteLine(value = -1)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True"]);
}

#[test]
fn gc_generation_for_object() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim obj As Object = New Byte(9) {}
        Dim gen As Integer = GC.GetGeneration(obj)
        Console.WriteLine(gen >= 0)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True"]);
}

#[test]
fn gc_get_generation_handles_null_as_minus_one() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim value As Integer = GC.GetGeneration(Nothing)
        Console.WriteLine(value = -1)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True"]);
}
