//! Overload resolution: `ref`/`out`/`params`, optional parameters, and named
//! arguments pick different targets at compile time.
use super::helpers::run_csharp;

#[test]
fn ref_overload_mutates_caller_storage_through_chosen_signature() {
    assert_eq!(
        run_csharp(
            r#"
void Scale(int value) { Console.WriteLine("byval:" + value); }
void Scale(ref int value) { value = value * 2; }
int n = 5;
Scale(ref n);
Console.WriteLine("after:" + n);
"#
        ),
        &["after:10"]
    );
}

#[test]
fn out_parameter_is_assigned_before_caller_observes_result() {
    assert_eq!(
        run_csharp(
            r#"
bool TryHalve(int input, out int half) {
    if (input % 2 != 0) {
        half = 0;
        return false;
    }
    half = input / 2;
    return true;
}
if (TryHalve(8, out var result)) {
    Console.WriteLine(result);
} else {
    Console.WriteLine("fail");
}
"#
        ),
        &["4"]
    );
}

#[test]
fn params_array_overload_receives_remaining_arguments_as_array() {
    assert_eq!(
        run_csharp(
            r#"
int Sum(params int[] values) {
    int total = 0;
    foreach (var v in values) total += v;
    return total;
}
Console.WriteLine(Sum(1, 2, 3));
Console.WriteLine(Sum());
"#
        ),
        &["6", "0"]
    );
}

#[test]
fn optional_parameter_uses_default_when_argument_omitted() {
    assert_eq!(
        run_csharp(
            r#"
string FormatLine(string text, int level = 1) {
    return level + ":" + text;
}
Console.WriteLine(FormatLine("ok"));
Console.WriteLine(FormatLine("warn", 3));
"#
        ),
        &["1:ok", "3:warn"]
    );
}

#[test]
fn named_arguments_can_reorder_optional_parameters_at_call_site() {
    assert_eq!(
        run_csharp(
            r#"
void Connect(string host, int port = 80, bool secure = false) {
    Console.WriteLine(host + ":" + port + ":" + secure);
}
Connect(secure: true, host: "api", port: 443);
"#
        ),
        &["api:443:True"]
    );
}

#[test]
fn more_specific_overload_wins_over_params_fallback_for_fixed_arity() {
    assert_eq!(
        run_csharp(
            r#"
string Describe(int value) { return "int:" + value; }
string Describe(params int[] values) { return "many:" + values.Length; }
Console.WriteLine(Describe(7));
Console.WriteLine(Describe(1, 2));
"#
        ),
        &["int:7", "many:2"]
    );
}
