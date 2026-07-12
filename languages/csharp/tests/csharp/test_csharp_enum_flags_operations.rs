//! `[Flags]` enums combine with bitwise ops; `HasFlag` tests membership.
use super::helpers::run_csharp;

#[test]
fn flags_enum_or_combines_independent_bits() {
    assert_eq!(
        run_csharp(
            r#"
[System.Flags]
enum Perm { None = 0, Read = 1, Write = 2 }
var value = Perm.Read | Perm.Write;
Console.WriteLine((int)value);
"#
        ),
        &["3"]
    );
}

#[test]
fn flags_enum_has_flag_detects_combined_permission() {
    assert_eq!(
        run_csharp(
            r#"
[System.Flags]
enum Perm { Read = 1, Write = 2, Execute = 4 }
var value = Perm.Read | Perm.Write;
Console.WriteLine(value.HasFlag(Perm.Write));
"#
        ),
        &["True"]
    );
}

#[test]
fn flags_enum_has_flag_reports_false_for_absent_bit() {
    assert_eq!(
        run_csharp(
            r#"
[System.Flags]
enum Perm { Read = 1, Write = 2, Execute = 4 }
var value = Perm.Read;
Console.WriteLine(value.HasFlag(Perm.Execute));
"#
        ),
        &["False"]
    );
}

#[test]
fn flags_enum_and_masks_to_intersection_of_bits() {
    assert_eq!(
        run_csharp(
            r#"
[System.Flags]
enum Perm { A = 1, B = 2, C = 4 }
var combined = Perm.A | Perm.B | Perm.C;
var masked = combined & Perm.B;
Console.WriteLine((int)masked);
"#
        ),
        &["2"]
    );
}

#[test]
fn flags_enum_xor_toggles_bits_present_in_one_operand() {
    assert_eq!(
        run_csharp(
            r#"
[System.Flags]
enum Perm { A = 1, B = 2 }
var value = (Perm.A | Perm.B) ^ Perm.A;
Console.WriteLine((int)value);
"#
        ),
        &["2"]
    );
}

#[test]
fn flags_enum_complement_within_byte_mask_inverts_bits() {
    assert_eq!(
        run_csharp(
            r#"
[System.Flags]
enum Perm : byte { A = 1, B = 2 }
var value = Perm.A | Perm.B;
var cleared = value & ~Perm.A;
Console.WriteLine((int)cleared);
"#
        ),
        &["2"]
    );
}

#[test]
fn plain_enum_cast_to_int_preserves_underlying_value() {
    assert_eq!(
        run_csharp(
            r#"
enum Level { Low = 1, Mid = 5, High = 9 }
Console.WriteLine((int)Level.Mid);
"#
        ),
        &["5"]
    );
}

#[test]
fn enum_parse_reads_name_into_typed_value() {
    assert_eq!(
        run_csharp(
            r#"
enum Color { Red, Green, Blue }
var value = System.Enum.Parse(typeof(Color), "Green");
Console.WriteLine(value);
"#
        ),
        &["Green"]
    );
}

#[test]
fn enum_to_string_returns_declared_identifier() {
    assert_eq!(
        run_csharp(
            r#"
enum Status { Idle, Running, Done }
Console.WriteLine(Status.Running.ToString());
"#
        ),
        &["Running"]
    );
}

#[test]
fn flags_enum_none_is_zero_and_or_identity() {
    assert_eq!(
        run_csharp(
            r#"
[System.Flags]
enum Perm { None = 0, Read = 1 }
var value = Perm.None | Perm.Read;
Console.WriteLine((int)value);
"#
        ),
        &["1"]
    );
}

#[test]
fn enum_switch_dispatches_on_underlying_constant_value() {
    assert_eq!(
        run_csharp(
            r#"
enum Mode { Alpha = 1, Beta = 2 }
string Label(Mode mode) {
    switch (mode) {
        case Mode.Alpha: return "a";
        case Mode.Beta: return "b";
        default: return "?";
    }
}
Console.WriteLine(Label(Mode.Beta));
"#
        ),
        &["b"]
    );
}
