//! `finally` interaction with `return`, nested `try`, and exception propagation —
//! control-flow contracts that are easy to get wrong in lowering.
use super::helpers::run_csharp;

#[test]
fn finally_runs_before_return_value_from_try_is_delivered_to_caller() {
    assert_eq!(
        run_csharp(
            r#"
int Pick() {
    try {
        return 2;
    } finally {
        Console.WriteLine("cleanup");
    }
}
Console.WriteLine(Pick());
"#
        ),
        &["cleanup", "2"]
    );
}

#[test]
fn finally_in_nested_try_runs_before_outer_catch_handles_exception() {
    assert_eq!(
        run_csharp(
            r#"
try {
    try {
        throw new Exception("boom");
    } finally {
        Console.WriteLine("inner-finally");
    }
} catch (Exception) {
    Console.WriteLine("outer-catch");
}
"#
        ),
        &["inner-finally", "outer-catch"]
    );
}

#[test]
fn catch_filter_when_clause_skips_handler_for_non_matching_predicate() {
    assert_eq!(
        run_csharp(
            r#"
string label = "start";
try {
    throw new Exception("code-404");
} catch (Exception e) when (e.Message.Contains("500")) {
    label = "wrong";
} catch (Exception e) when (e.Message.Contains("404")) {
    label = "matched";
}
Console.WriteLine(label);
"#
        ),
        &["matched"]
    );
}

#[test]
fn exception_rethrown_from_catch_is_handled_by_enclosing_try() {
    assert_eq!(
        run_csharp(
            r#"
string trace = "";
try {
    try {
        throw new Exception("first");
    } catch (Exception) {
        trace += "inner;";
        throw new Exception("second");
    }
} catch (Exception e) {
    trace += "outer:" + e.Message;
}
Console.WriteLine(trace);
"#
        ),
        &["inner;outer:second"]
    );
}

#[test]
fn try_without_catch_still_executes_finally_when_body_throws() {
    assert_eq!(
        run_csharp(
            r#"
string trace = "";
try {
    try {
        throw new Exception("fail");
    } finally {
        trace += "finally;";
    }
} catch (Exception) {
    trace += "handled;";
}
Console.WriteLine(trace);
"#
        ),
        &["finally;handled;"]
    );
}
