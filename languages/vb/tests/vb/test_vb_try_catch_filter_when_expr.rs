use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Catch ... When Conditional Exception Filters
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_catch_when_filter_evaluated_true() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Try
            Dim errCode As Integer = 404
            Throw New Exception("Page Not Found")
        Catch ex As Exception When ex.Message.Contains("404") OrElse True
            Console.WriteLine("Filtered Catch: " & ex.Message)
        Catch ex As Exception
            Console.WriteLine("General Catch")
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Filtered Catch: Page Not Found"]);
}

#[test]
fn test_vb_catch_when_filter_evaluated_false_falls_through() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Try
            Throw New InvalidOperationException("Operation Failed")
        Catch ex As Exception When ex.Message.Contains("Database")
            Console.WriteLine("Database Catch")
        Catch ex As Exception
            Console.WriteLine("Fallback Catch: " & ex.GetType().Name)
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["Fallback Catch: InvalidOperationException"]
    );
}

#[test]
fn test_vb_catch_when_side_effects_in_filter() {
    let src = r#"
Imports System

Module Program
    Public FilterCount As Integer = 0

    Public Function LogAndCheck(ex As Exception) As Boolean
        FilterCount += 1
        Return False
    End Function

    Sub Main()
        Try
            Throw New Exception("Test")
        Catch ex As Exception When LogAndCheck(ex)
            Console.WriteLine("Caught in filter")
        Catch ex As Exception
            Console.WriteLine("Caught in fallback, count=" & FilterCount)
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Caught in fallback, count=1"]);
}
