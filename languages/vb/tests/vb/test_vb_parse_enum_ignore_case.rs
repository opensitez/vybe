use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: System.Enum Parsing, Case Insensitivity & Formatting
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_enum_parse_case_sensitive() {
    let src = r#"
Imports System

Enum Status
    Pending
    Active
End Enum

Module Program
    Sub Main()
        Dim s As Status = CType([Enum].Parse(GetType(Status), "Active"), Status)
        Console.WriteLine(s.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Active"]);
}

#[test]
fn test_vb_enum_parse_ignore_case() {
    let src = r#"
Imports System

Enum Command
    Start
    Stop
End Enum

Module Program
    Sub Main()
        ' Enum.Parse(type, string, ignoreCase:=True)
        Dim cmd As Command = CType([Enum].Parse(GetType(Command), "start", True), Command)
        Console.WriteLine(cmd.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Start"]);
}

#[test]
fn test_vb_enum_try_parse_generic_ignore_case() {
    let src = r#"
Imports System

Enum LogLevel
    Debug
    Warning
    Error
End Enum

Module Program
    Sub Main()
        Dim level As LogLevel
        Dim ok = [Enum].TryParse(Of LogLevel)("warning", True, level)
        Console.WriteLine(ok & "|" & level.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|Warning"]);
}

#[test]
fn test_vb_enum_parse_numeric_string() {
    let src = r#"
Imports System

Enum Priority
    Low = 1
    Medium = 2
    High = 3
End Enum

Module Program
    Sub Main()
        Dim p As Priority = CType([Enum].Parse(GetType(Priority), "2"), Priority)
        Console.WriteLine(p.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Medium"]);
}

#[test]
fn test_vb_enum_flags_parse_comma_separated_string() {
    let src = r#"
Imports System

<Flags>
Enum Permissions
    Read = 1
    Write = 2
    Execute = 4
End Enum

Module Program
    Sub Main()
        Dim p As Permissions = CType([Enum].Parse(GetType(Permissions), "Read, Write"), Permissions)
        Console.WriteLine(p.HasFlag(Permissions.Read) & "|" & p.HasFlag(Permissions.Write) & "|" & p.HasFlag(Permissions.Execute))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True|False"]);
}

#[test]
fn test_vb_enum_is_defined_check() {
    let src = r#"
Imports System

Enum Colors
    Red = 1
    Green = 2
End Enum

Module Program
    Sub Main()
        Console.WriteLine([Enum].IsDefined(GetType(Colors), "Red") & "|" & [Enum].IsDefined(GetType(Colors), 99))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|False"]);
}

#[test]
fn test_vb_enum_get_names_and_get_values() {
    let src = r#"
Imports System

Enum Days
    Mon = 1
    Tue = 2
End Enum

Module Program
    Sub Main()
        Dim names = [Enum].GetNames(GetType(Days))
        Dim values = [Enum].GetValues(GetType(Days))
        Console.WriteLine(String.Join(",", names) & "|" & values.Length)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Mon,Tue|2"]);
}

#[test]
fn test_vb_enum_get_underlying_type() {
    let src = r#"
Imports System

Enum ShortEnum As Short
    A = 10
End Enum

Module Program
    Sub Main()
        Dim t = [Enum].GetUnderlyingType(GetType(ShortEnum))
        Console.WriteLine(t.Name)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Int16"]);
}

#[test]
fn test_vb_enum_format_specifiers_g_f_d_x() {
    let src = r#"
Imports System

<Flags>
Enum Modes
    Read = 1
    Write = 2
End Enum

Module Program
    Sub Main()
        Dim m = Modes.Read Or Modes.Write
        Console.WriteLine([Enum].Format(GetType(Modes), m, "G") & "|" & [Enum].Format(GetType(Modes), m, "D") & "|" & [Enum].Format(GetType(Modes), m, "X"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Read, Write|3|00000003"]);
}

#[test]
fn test_vb_enum_parse_invalid_string_throws_argument_exception() {
    let src = r#"
Imports System

Enum State
    OnState
    OffState
End Enum

Module Program
    Sub Main()
        Try
            [Enum].Parse(GetType(State), "InvalidStateValue")
        Catch ex As ArgumentException
            Console.WriteLine("ArgumentException Caught on Invalid Enum Name")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["ArgumentException Caught on Invalid Enum Name"]
    );
}

#[test]
fn test_vb_enum_try_parse_invalid_returns_false() {
    let src = r#"
Imports System

Enum Direction
    North
    South
End Enum

Module Program
    Sub Main()
        Dim dir As Direction
        Dim ok = [Enum].TryParse(Of Direction)("West", dir)
        Console.WriteLine(ok & "|" & dir)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["False|North"]);
}

#[test]
fn test_vb_enum_to_object_conversion() {
    let src = r#"
Imports System

Enum Suit
    Hearts = 1
    Spades = 2
End Enum

Module Program
    Sub Main()
        Dim obj = [Enum].ToObject(GetType(Suit), 2)
        Console.WriteLine(obj.ToString() & "|" & obj.GetType().Name)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Spades|Suit"]);
}

#[test]
fn test_vb_enum_has_flag_method() {
    let src = r#"
Imports System

<Flags>
Enum Attributes
    None = 0
    ReadOnly = 1
    Hidden = 2
    System = 4
End Enum

Module Program
    Sub Main()
        Dim attr = Attributes.ReadOnly Or Attributes.Hidden
        Console.WriteLine(attr.HasFlag(Attributes.ReadOnly) & "|" & attr.HasFlag(Attributes.System))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|False"]);
}

#[test]
fn test_vb_enum_bitwise_and_or_xor_operators() {
    let src = r#"
Imports System

<Flags>
Enum Options
    OptA = 1
    OptB = 2
    OptC = 4
End Enum

Module Program
    Sub Main()
        Dim combined = Options.OptA Or Options.OptB
        Dim toggled = combined Xor Options.OptA
        Console.WriteLine(toggled.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["OptB"]);
}

#[test]
fn test_vb_enum_comparison_operators() {
    let src = r#"
Enum Level
    Low = 10
    High = 20
End Enum

Module Program
    Sub Main()
        Dim l1 = Level.Low
        Dim l2 = Level.High
        Console.WriteLine((l1 < l2) & "|" & (l1 = Level.Low))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True"]);
}

#[test]
fn test_vb_enum_parse_undefined_numeric_value() {
    let src = r#"
Imports System

Enum Code
    Valid = 100
End Enum

Module Program
    Sub Main()
        ' Enum.Parse on undefined numeric string "999" succeeds and returns underlying value!
        Dim c As Code = CType([Enum].Parse(GetType(Code), "999"), Code)
        Console.WriteLine(CInt(c))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["999"]);
}

#[test]
fn test_vb_enum_parse_null_string_throws_argument_null() {
    let src = r#"
Imports System

Enum Category
    CatA
End Enum

Module Program
    Sub Main()
        Try
            [Enum].Parse(GetType(Category), Nothing)
        Catch ex As ArgumentNullException
            Console.WriteLine("ArgumentNullException Caught on Null Enum Parse")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["ArgumentNullException Caught on Null Enum Parse"]
    );
}

#[test]
fn test_vb_enum_get_names_sorted_by_underlying_value() {
    let src = r#"
Imports System

Enum UnorderedValues
    Second = 20
    First = 10
End Enum

Module Program
    Sub Main()
        Dim names = [Enum].GetNames(GetType(UnorderedValues))
        Console.WriteLine(String.Join(",", names))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["First,Second"]);
}

#[test]
fn test_vb_enum_ulong_underlying_type_parsing() {
    let src = r#"
Imports System

Enum BigEnum As ULong
    MaxVal = 18000000000000000000UL
End Enum

Module Program
    Sub Main()
        Dim e As BigEnum = CType([Enum].Parse(GetType(BigEnum), "MaxVal"), BigEnum)
        Console.WriteLine(CULng(e))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["18000000000000000000"]);
}

#[test]
fn test_vb_enum_parse_whitespace_trimmed_automatically() {
    let src = r#"
Imports System

Enum OptionType
    Enabled
End Enum

Module Program
    Sub Main()
        Dim opt As OptionType = CType([Enum].Parse(GetType(OptionType), "  Enabled  "), OptionType)
        Console.WriteLine(opt.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Enabled"]);
}
