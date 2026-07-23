use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Named Tuples & Nested Structural Combinations
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_named_tuple_explicit_field_names() {
    let src = r#"
Module Program
    Sub Main()
        Dim pt As (X As Integer, Y As Integer) = (10, 20)
        Console.WriteLine(pt.X & "," & pt.Y)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10,20"]);
}

#[test]
fn test_vb_named_tuple_positional_item_access_compatibility() {
    let src = r#"
Module Program
    Sub Main()
        Dim pt As (X As Integer, Y As Integer) = (10, 20)
        ' Positional .Item1 and .Item2 are aliases for .X and .Y
        Console.WriteLine(pt.Item1 & "," & pt.Item2)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10,20"]);
}

#[test]
fn test_vb_nested_tuple_structure_access() {
    let src = r#"
Module Program
    Sub Main()
        Dim line As (StartPt As (X As Integer, Y As Integer), EndPt As (X As Integer, Y As Integer))
        line.StartPt = (0, 0)
        line.EndPt = (10, 20)
        Console.WriteLine(line.StartPt.X & "," & line.StartPt.Y & " -> " & line.EndPt.X & "," & line.EndPt.Y)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0,0 -> 10,20"]);
}

#[test]
fn test_vb_named_tuple_assignment_name_erasure_compatibility() {
    let src = r#"
Module Program
    Sub Main()
        Dim t1 As (Width As Integer, Height As Integer) = (100, 200)
        Dim t2 As (W As Integer, H As Integer) = t1 ' Names are erased at runtime; underlying ValueTuple is compatible!
        Console.WriteLine(t2.W & "x" & t2.H)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["100x200"]);
}

#[test]
fn test_vb_named_tuple_method_parameters() {
    let src = r#"
Module Program
    Private Function FormatPerson(p As (Name As String, Age As Integer)) As String
        Return p.Name & " (" & p.Age & ")"
    End Function

    Sub Main()
        Console.WriteLine(FormatPerson(("Alice", 30)))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Alice (30)"]);
}

#[test]
fn test_vb_named_tuple_optional_method_parameters() {
    let src = r#"
Module Program
    Private Function GetConfig(Optional cfg As (Host As String, Port As Integer) = Nothing) As String
        If cfg.Host Is Nothing Then Return "default:8080"
        Return cfg.Host & ":" & cfg.Port
    End Function

    Sub Main()
        Console.WriteLine(GetConfig() & "|" & GetConfig(("localhost", 9000)))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["default:8080|localhost:9000"]);
}

#[test]
fn test_vb_named_tuple_dictionary_as_value() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New Dictionary(Of String, (Min As Integer, Max As Integer))()
        dict("Age") = (18, 65)
        Console.WriteLine(dict("Age").Min & " To " & dict("Age").Max)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["18 To 65"]);
}

#[test]
fn test_vb_named_tuple_dictionary_as_key() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New Dictionary(Of (X As Integer, Y As Integer), String)()
        dict((1, 2)) = "Top"
        Console.WriteLine(dict((1, 2)))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Top"]);
}

#[test]
fn test_vb_named_tuple_list_sorting_by_tuple_element() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim items As New List(Of (Name As String, Priority As Integer)) From {
            ("TaskB", 2),
            ("TaskA", 1),
            ("TaskC", 3)
        }
        items.Sort(Function(x, y) x.Priority.CompareTo(y.Priority))
        For Each item In items
            Console.WriteLine(item.Name & ":" & item.Priority)
        Next
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["TaskA:1", "TaskB:2", "TaskC:3"]);
}

#[test]
fn test_vb_named_tuple_generic_class_field() {
    let src = r#"
Class StateHolder(Of T)
    Public CurrentState As (Status As String, Data As T)
    Public Sub New(s As String, d As T)
        CurrentState = (s, d)
    End Sub
End Class

Module Program
    Sub Main()
        Dim sh As New StateHolder(Of Integer)("OK", 200)
        Console.WriteLine(sh.CurrentState.Status & "=" & sh.CurrentState.Data)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["OK=200"]);
}

#[test]
fn test_vb_named_tuple_interface_property() {
    let src = r#"
Interface IScheduledTask
    ReadOnly Property ScheduleInfo As (StartTime As String, DurationMinutes As Integer)
End Interface

Class MaintenanceTask
    Implements IScheduledTask
    Public ReadOnly Property ScheduleInfo As (StartTime As String, DurationMinutes As Integer) Implements IScheduledTask.ScheduleInfo
        Get
            Return ("02:00", 60)
        End Get
    End Property
End Class

Module Program
    Sub Main()
        Dim task As IScheduledTask = New MaintenanceTask()
        Console.WriteLine(task.ScheduleInfo.StartTime & " for " & task.ScheduleInfo.DurationMinutes & "m")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["02:00 for 60m"]);
}

#[test]
fn test_vb_named_tuple_in_event_args() {
    let src = r#"
Imports System

Class TaskRunner
    Public Event TaskProgress As Action(Of (StepName As String, Percent As Integer))
    Public Sub Report(name As String, pct As Integer)
        RaiseEvent TaskProgress((name, pct))
    End Sub
End Class

Module Program
    Sub Main()
        Dim runner As New TaskRunner()
        AddHandler runner.TaskProgress, Sub(info) Console.WriteLine(info.StepName & " " & info.Percent & "%")
        runner.Report("Download", 75)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Download 75%"]);
}

#[test]
fn test_vb_named_tuple_field_mutation() {
    let src = r#"
Module Program
    Sub Main()
        Dim mutableTuple As (Count As Integer, Label As String) = (0, "Init")
        mutableTuple.Count += 5
        mutableTuple.Label = "Updated"
        Console.WriteLine(mutableTuple.Count & "|" & mutableTuple.Label)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["5|Updated"]);
}

#[test]
fn test_vb_named_tuple_conversion_from_value_tuple() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim rawTuple As ValueTuple(Of String, Integer) = ValueTuple.Create("Raw", 99)
        Dim namedTuple As (Tag As String, Val As Integer) = rawTuple
        Console.WriteLine(namedTuple.Tag & "=" & namedTuple.Val)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Raw=99"]);
}

#[test]
fn test_vb_named_tuple_conversion_to_value_tuple() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim namedTuple As (Tag As String, Val As Integer) = ("Named", 42)
        Dim rawTuple As ValueTuple(Of String, Integer) = namedTuple
        Console.WriteLine(rawTuple.Item1 & ":" & rawTuple.Item2)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Named:42"]);
}

#[test]
fn test_vb_named_tuple_linq_select_projection() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim nums = {1, 2, 3}
        Dim tuples = From n In nums Select Entry = (Value:=n, Square:=n * n)
        For Each t In tuples
            Console.WriteLine(t.Value & "^2=" & t.Square)
        Next
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1^2=1", "2^2=4", "3^2=9"]);
}

#[test]
fn test_vb_named_tuple_8_elements_large_tuple() {
    let src = r#"
Module Program
    Sub Main()
        Dim t = (A:=1, B:=2, C:=3, D:=4, E:=5, F:=6, G:=7, H:=8)
        Console.WriteLine(t.A & "+" & t.H)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1+8"]);
}

#[test]
fn test_vb_named_tuple_struct_inside_tuple() {
    let src = r#"
Structure Point
    Public X As Integer
    Public Y As Integer
    Public Sub New(x As Integer, y As Integer) : Me.X = x : Me.Y = y : End Sub
End Structure

Module Program
    Sub Main()
        Dim data As (Location As Point, Name As String) = (New Point(5, 10), "Target")
        Console.WriteLine(data.Name & " at " & data.Location.X & "," & data.Location.Y)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Target at 5,10"]);
}

#[test]
fn test_vb_named_tuple_enum_field() {
    let src = r#"
Enum Level
    Info
    Warning
    ErrorVal
End Enum

Module Program
    Sub Main()
        Dim log As (Severity As Level, Message As String) = (Level.Warning, "Disk Space Low")
        Console.WriteLine("[" & log.Severity.ToString() & "] " & log.Message)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["[Warning] Disk Space Low"]);
}

#[test]
fn test_vb_named_tuple_nested_array_of_tuples() {
    let src = r#"
Module Program
    Sub Main()
        Dim grid As (Row As Integer, Cols As (ColName As String, Val As Integer)())() = {
            (1, {("C1", 10), ("C2", 20)})
        }
        Console.WriteLine(grid(0).Row & "->" & grid(0).Cols(0).ColName & "=" & grid(0).Cols(0).Val)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1->C1=10"]);
}
