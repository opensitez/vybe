//! `switch` with type and constant patterns — first match wins.
use super::helpers::run_csharp;

#[test]
fn switch_on_string_matches_exact_literal_case() {
    assert_eq!(
        run_csharp(
            r#"
string Pick(string key) {
    switch (key) {
        case "go": return "G";
        case "stop": return "S";
        default: return "?";
    }
}
Console.WriteLine(Pick("go"));
"#
        ),
        &["G"]
    );
}

#[test]
fn switch_on_int_falls_through_to_default_when_no_case_matches() {
    assert_eq!(
        run_csharp(
            r#"
int code = 99;
string label = "";
switch (code) {
    case 1: label = "one"; break;
    default: label = "other"; break;
}
Console.WriteLine(label);
"#
        ),
        &["other"]
    );
}

#[test]
fn switch_expression_returns_value_from_matching_arm() {
    assert_eq!(
        run_csharp(
            r#"
int n = 2;
string word = n switch { 1 => "one", 2 => "two", _ => "many" };
Console.WriteLine(word);
"#
        ),
        &["two"]
    );
}

#[test]
fn is_int_pattern_binds_variable_in_true_branch() {
    assert_eq!(
        run_csharp(
            r#"
object boxed = 12;
if (boxed is int value) {
    Console.WriteLine(value + 1);
} else {
    Console.WriteLine("no");
}
"#
        ),
        &["13"]
    );
}

#[test]
fn is_string_pattern_fails_for_non_matching_runtime_type() {
    assert_eq!(
        run_csharp(
            r#"
object boxed = 12;
if (boxed is string text) {
    Console.WriteLine(text);
} else {
    Console.WriteLine("not-string");
}
"#
        ),
        &["not-string"]
    );
}

#[test]
fn switch_statement_with_when_clause_filters_case() {
    assert_eq!(
        run_csharp(
            r#"
int n = 8;
string size = n switch {
    < 0 => "neg",
    >= 0 and < 10 => "small",
    _ => "big"
};
Console.WriteLine(size);
"#
        ),
        &["small"]
    );
}

#[test]
fn switch_on_enum_uses_symbolic_case_labels() {
    assert_eq!(
        run_csharp(
            r#"
enum Tier { Free, Pro }
string Name(Tier tier) => tier switch {
    Tier.Free => "free",
    Tier.Pro => "pro",
    _ => "unknown"
};
Console.WriteLine(Name(Tier.Pro));
"#
        ),
        &["pro"]
    );
}

#[test]
fn is_pattern_with_null_constant_detects_null_reference() {
    assert_eq!(
        run_csharp(
            r#"
string text = null;
Console.WriteLine(text is null);
"#
        ),
        &["True"]
    );
}

#[test]
fn is_not_pattern_negates_type_test() {
    assert_eq!(
        run_csharp(
            r#"
object value = 3.14;
Console.WriteLine(value is not int);
"#
        ),
        &["True"]
    );
}

#[test]
fn switch_on_bool_has_separate_true_and_false_arms() {
    assert_eq!(
        run_csharp(
            r#"
bool ok = false;
string label = ok switch { true => "yes", false => "no" };
Console.WriteLine(label);
"#
        ),
        &["no"]
    );
}

#[test]
fn relational_pattern_greater_or_equal_matches_boundary_value() {
    assert_eq!(
        run_csharp(
            r#"
int score = 100;
string grade = score switch {
    >= 90 => "A",
    >= 80 => "B",
    _ => "C"
};
Console.WriteLine(grade);
"#
        ),
        &["A"]
    );
}

#[test]
fn switch_nested_inside_loop_accumulates_labels_per_iteration() {
    assert_eq!(
        run_csharp(
            r#"
string trace = "";
for (int i = 0; i < 3; i++) {
    trace += i switch { 0 => "a", 1 => "b", _ => "c" };
}
Console.WriteLine(trace);
"#
        ),
        &["abc"]
    );
}
