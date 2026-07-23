use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: GoTo Statement, Labels & Line Jump Control Flow
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_goto_forward_jump() {
    let src = r#"
Module Program
    Sub Main()
        Console.WriteLine("Start")
        GoTo TargetLabel
        Console.WriteLine("Skipped")
TargetLabel:
        Console.WriteLine("End")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Start", "End"]);
}

#[test]
fn test_vb_goto_backward_loop_jump() {
    let src = r#"
Module Program
    Sub Main()
        Dim count = 0
StartLoop:
        count += 1
        If count < 3 Then GoTo StartLoop
        Console.WriteLine(count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3"]);
}

#[test]
fn test_vb_goto_exit_nested_loops() {
    let src = r#"
Module Program
    Sub Main()
        For r As Integer = 1 To 5
            For c As Integer = 1 To 5
                If r = 2 AndAlso c = 2 Then GoTo BreakAll
            Next
        Next
BreakAll:
        Console.WriteLine("Broke Out of All Loops")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Broke Out of All Loops"]);
}

#[test]
fn test_vb_goto_numeric_line_labels_legacy() {
    let src = r#"
Module Program
    Sub Main()
10:     Console.WriteLine("Line 10")
        GoTo 30
20:     Console.WriteLine("Line 20")
30:     Console.WriteLine("Line 30")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Line 10", "Line 30"]);
}

#[test]
fn test_vb_goto_inside_if_statement_branch() {
    let src = r#"
Module Program
    Sub Main()
        Dim flag = True
        If flag Then
            GoTo SuccessLabel
        Else
            GoTo FailureLabel
        End If
SuccessLabel:
        Console.WriteLine("Success Branch")
        Exit Sub
FailureLabel:
        Console.WriteLine("Failure Branch")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["SuccessBranch"]);
}

#[test]
fn test_vb_goto_inside_select_case_branch() {
    let src = r#"
Module Program
    Sub Main()
        Dim mode = 2
        Select Case mode
            Case 1
                Console.WriteLine("Mode 1")
            Case 2
                GoTo SpecialMode
        End Select
        Console.WriteLine("Normal End")
        Exit Sub
SpecialMode:
        Console.WriteLine("Special Mode Jumped")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Special Mode Jumped"]);
}

#[test]
fn test_vb_goto_out_of_try_block() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Try
            Console.WriteLine("In Try Block")
            GoTo ExternalLabel
        Catch ex As Exception
            Console.WriteLine("In Catch")
        Finally
            Console.WriteLine("In Finally")
        End Try
ExternalLabel:
        Console.WriteLine("Outside Try")
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["In Try Block", "In Finally", "Outside Try"]
    );
}

#[test]
fn test_vb_goto_out_of_catch_block() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Try
            Throw New InvalidOperationException("Fail")
        Catch ex As InvalidOperationException
            Console.WriteLine("In Catch")
            GoTo HandledLabel
        Finally
            Console.WriteLine("In Finally")
        End Try
HandledLabel:
        Console.WriteLine("Handled External")
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["In Catch", "In Finally", "Handled External"]
    );
}

#[test]
fn test_vb_goto_multiple_forward_jumps() {
    let src = r#"
Module Program
    Sub Main()
        GoTo L1
L2:
        Console.WriteLine("Step 2")
        GoTo L3
L1:
        Console.WriteLine("Step 1")
        GoTo L2
L3:
        Console.WriteLine("Step 3")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Step 1", "Step 2", "Step 3"]);
}

#[test]
fn test_vb_goto_label_case_insensitivity() {
    let src = r#"
Module Program
    Sub Main()
        GoTo mycustomlabel
MYCUSTOMLABEL:
        Console.WriteLine("Label Matched Case Insensitively")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Label Matched Case Insensitively"]);
}

#[test]
fn test_vb_goto_label_scoped_to_procedure() {
    let src = r#"
Module Program
    Private Sub ProcA()
        GoTo CommonLabel
CommonLabel:
        Console.WriteLine("ProcA Label")
    End Sub

    Private Sub ProcB()
        GoTo CommonLabel
CommonLabel:
        Console.WriteLine("ProcB Label")
    End Sub

    Sub Main()
        ProcA()
        ProcB()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["ProcA Label", "ProcB Label"]);
}

#[test]
fn test_vb_goto_state_machine_simulation() {
    let src = r#"
Module Program
    Sub Main()
        Dim state = 0
State0:
        Console.WriteLine("State0")
        state = 1
        GoTo State1
State1:
        Console.WriteLine("State1")
        state = 2
        If state = 2 Then GoTo StateEnd
StateEnd:
        Console.WriteLine("StateEnd")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["State0", "State1", "StateEnd"]);
}

#[test]
fn test_vb_goto_inside_using_block_disposes_resource() {
    let src = r#"
Imports System
Imports System.IO

Module Program
    Sub Main()
        Using ms As New MemoryStream()
            ms.WriteByte(1)
            GoTo ExitUsing
        End Using
ExitUsing:
        Console.WriteLine("Exited Using Block via GoTo")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Exited Using Block via GoTo"]);
}

#[test]
fn test_vb_goto_out_of_for_each_loop() {
    let src = r#"
Module Program
    Sub Main()
        Dim items As String() = {"A", "B", "C"}
        For Each item In items
            If item = "B" Then GoTo FoundB
        Next
        Exit Sub
FoundB:
        Console.WriteLine("Found B via GoTo")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Found B via GoTo"]);
}

#[test]
fn test_vb_goto_out_of_do_while_loop() {
    let src = r#"
Module Program
    Sub Main()
        Dim i = 0
        Do While True
            i += 1
            If i = 5 Then GoTo LoopExit
        Loop
LoopExit:
        Console.WriteLine("Exited Do Loop: " & i)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Exited Do Loop: 5"]);
}

#[test]
fn test_vb_goto_backward_counter_accumulation() {
    let src = r#"
Module Program
    Sub Main()
        Dim total = 0
        Dim i = 1
Accumulate:
        total += i
        i += 1
        If i <= 5 Then GoTo Accumulate
        Console.WriteLine(total)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["15"]);
}

#[test]
fn test_vb_goto_label_following_end_sub() {
    let src = r#"
Module Program
    Private Function TestFunc() As String
        GoTo ResLabel
        Return "Unreachable"
ResLabel:
        Return "FunctionResult"
    End Function

    Sub Main()
        Console.WriteLine(TestFunc())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["FunctionResult"]);
}

#[test]
fn test_vb_goto_label_with_same_name_as_variable() {
    let src = r#"
Module Program
    Sub Main()
        Dim Target As String = "VarVal"
        GoTo Target
        Console.WriteLine("Skipped")
Target:
        Console.WriteLine(Target)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["VarVal"]);
}

#[test]
fn test_vb_goto_clean_exit_sequence() {
    let src = r#"
Module Program
    Sub Main()
        Dim success = True
        If Not success Then GoTo Cleanup
        Console.WriteLine("Processing...")
Cleanup:
        Console.WriteLine("Cleanup Done")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Processing...", "Cleanup Done"]);
}

#[test]
fn test_vb_goto_label_with_trailing_colon() {
    let src = r#"
Module Program
    Sub Main()
        GoTo LabelWithColon
LabelWithColon:
        Console.WriteLine("Label Passed")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Label Passed"]);
}
