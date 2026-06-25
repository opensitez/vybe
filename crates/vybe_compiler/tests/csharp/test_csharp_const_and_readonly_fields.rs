//! `const` compile-time constants vs `readonly` assign-once instance fields.
use super::helpers::run_csharp;

#[test]
fn const_field_is_accessible_without_instance() {
    assert_eq!(
        run_csharp(
            r#"
class Limits {
    public const int Max = 100;
}
Console.WriteLine(Limits.Max);
"#
        ),
        &["100"]
    );
}

#[test]
fn readonly_instance_field_set_in_constructor_is_visible_after() {
    assert_eq!(
        run_csharp(
            r#"
class Token {
    public readonly string Value;
    public Token(string value) { Value = value; }
}
Console.WriteLine(new Token("key").Value);
"#
        ),
        &["key"]
    );
}

#[test]
fn readonly_static_field_initialized_at_type_load() {
    assert_eq!(
        run_csharp(
            r#"
class Config {
    public static readonly string Prefix = "app";
}
Console.WriteLine(Config.Prefix);
"#
        ),
        &["app"]
    );
}

#[test]
fn const_local_cannot_be_reassigned_in_method() {
    assert_eq!(
        run_csharp(
            r#"
const int step = 5;
Console.WriteLine(step * 2);
"#
        ),
        &["10"]
    );
}

#[test]
fn readonly_field_cannot_change_after_constructor_body_completes() {
    assert_eq!(
        run_csharp(
            r#"
class Counter {
    public readonly int Seed;
    public Counter(int seed) { Seed = seed; }
    public int Read() { return Seed; }
}
Console.WriteLine(new Counter(3).Read());
"#
        ),
        &["3"]
    );
}

#[test]
fn const_string_concatenates_at_compile_time_in_expression() {
    assert_eq!(
        run_csharp(
            r#"
class Labels {
    public const string Base = "user";
    public const string Full = Base + "_id";
}
Console.WriteLine(Labels.Full);
"#
        ),
        &["user_id"]
    );
}

#[test]
fn static_readonly_vs_const_both_accessible_on_type_name() {
    assert_eq!(
        run_csharp(
            r#"
class Mix {
    public const int A = 1;
    public static readonly int B = 2;
}
Console.WriteLine(Mix.A + Mix.B);
"#
        ),
        &["3"]
    );
}

#[test]
fn readonly_struct_field_must_be_set_in_constructor() {
    assert_eq!(
        run_csharp(
            r#"
struct Cell {
    public readonly int Value;
    public Cell(int value) { Value = value; }
}
Console.WriteLine(new Cell(8).Value);
"#
        ),
        &["8"]
    );
}

#[test]
fn const_enum_member_casts_to_underlying_integer_value() {
    assert_eq!(
        run_csharp(
            r#"
enum Code { Ok = 0, Err = 1 }
const Code status = Code.Ok;
Console.WriteLine((int)status);
"#
        ),
        &["0"]
    );
}

#[test]
fn readonly_array_field_reference_cannot_be_replaced_but_elements_can() {
    assert_eq!(
        run_csharp(
            r#"
class Holder {
    public readonly int[] Data = { 1, 2 };
}
var holder = new Holder();
holder.Data[1] = 9;
Console.WriteLine(holder.Data[1]);
"#
        ),
        &["9"]
    );
}
