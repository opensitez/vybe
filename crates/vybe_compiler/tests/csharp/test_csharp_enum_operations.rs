//! Enum declaration, parsing, formatting, flags, and iteration.
use super::helpers::run_csharp;

#[test]
fn enum_value_assigned_and_compared() {
    assert_eq!(
        run_csharp(r#"enum Color{Red,Green,Blue} var c=Color.Green; Console.WriteLine(c==Color.Green);"#),
        &["True"]
    );
}

#[test]
fn enum_tostring_returns_member_name() {
    assert_eq!(
        run_csharp(r#"enum Status{Pending,Active,Done} Console.WriteLine(Status.Active.ToString());"#),
        &["Active"]
    );
}

#[test]
fn enum_cast_to_underlying_int_type() {
    assert_eq!(
        run_csharp(r#"enum Priority{Low=1,Medium=2,High=3} Console.WriteLine((int)Priority.High);"#),
        &["3"]
    );
}

#[test]
fn enum_parse_converts_string_name_to_value() {
    assert_eq!(
        run_csharp(
            r#"enum Day{Mon,Tue,Wed,Thu,Fri}
var d = (Day)System.Enum.Parse(typeof(Day),"Wed");
Console.WriteLine(d);"#
        ),
        &["Wed"]
    );
}

#[test]
fn enum_get_names_returns_all_member_names() {
    assert_eq!(
        run_csharp(
            r#"enum Coin{Penny,Nickel,Dime}
Console.WriteLine(System.Enum.GetNames(typeof(Coin)).Length);"#
        ),
        &["3"]
    );
}

#[test]
fn flags_enum_has_flag_detects_combined_bit() {
    assert_eq!(
        run_csharp(
            r#"[System.Flags] enum Perm{None=0,Read=1,Write=2,Execute=4}
var p = Perm.Read | Perm.Write;
Console.WriteLine(p.HasFlag(Perm.Read));
Console.WriteLine(p.HasFlag(Perm.Execute));"#
        ),
        &["True", "False"]
    );
}

#[test]
fn flags_enum_combined_value_has_expected_integer() {
    assert_eq!(
        run_csharp(
            r#"[System.Flags] enum Perm{None=0,Read=1,Write=2,Execute=4}
Console.WriteLine((int)(Perm.Read|Perm.Execute));"#
        ),
        &["5"]
    );
}

#[test]
fn enum_is_defined_returns_false_for_out_of_range_int() {
    assert_eq!(
        run_csharp(
            r#"enum Level{Low=0,Mid=1,High=2}
Console.WriteLine(System.Enum.IsDefined(typeof(Level), 99));"#
        ),
        &["False"]
    );
}

#[test]
fn enum_with_explicit_underlying_byte_type() {
    assert_eq!(
        run_csharp(
            r#"enum Small:byte{A=1,B=200}
Console.WriteLine((byte)Small.B);"#
        ),
        &["200"]
    );
}
