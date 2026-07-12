//! Exception filters (`catch (E e) when (...)`) select handler at throw site.
use super::helpers::run_csharp;

#[test]
fn catch_when_filter_matches_specific_message_content() {
    assert_eq!(
        run_csharp(
            r#"
try {
    throw new System.Exception("code=404");
} catch (System.Exception e) when (e.Message.Contains("404")) {
    Console.WriteLine("not found");
} catch (System.Exception) {
    Console.WriteLine("other");
}
"#
        ),
        &["not found"]
    );
}

#[test]
fn catch_when_filter_falls_through_to_next_handler_when_false() {
    assert_eq!(
        run_csharp(
            r#"
try {
    throw new System.Exception("code=500");
} catch (System.Exception e) when (e.Message.Contains("404")) {
    Console.WriteLine("not found");
} catch (System.Exception) {
    Console.WriteLine("server error");
}
"#
        ),
        &["server error"]
    );
}

#[test]
fn catch_when_filter_can_evaluate_arbitrary_boolean_expression() {
    assert_eq!(
        run_csharp(
            r#"
int threshold = 10;
try {
    throw new System.InvalidOperationException("value=15");
} catch (System.InvalidOperationException e) when (threshold < 20) {
    Console.WriteLine("caught with threshold");
}
"#
        ),
        &["caught with threshold"]
    );
}

#[test]
fn rethrow_preserves_original_stack_trace() {
    assert_eq!(
        run_csharp(
            r#"
string result = "";
try {
    try {
        throw new System.Exception("original");
    } catch (System.Exception) {
        throw;
    }
} catch (System.Exception e) {
    result = e.Message;
}
Console.WriteLine(result);
"#
        ),
        &["original"]
    );
}

#[test]
fn inner_exception_chains_to_outer_exception_cause() {
    assert_eq!(
        run_csharp(
            r#"
try {
    try {
        throw new System.Exception("root cause");
    } catch (System.Exception inner) {
        throw new System.InvalidOperationException("wrapped", inner);
    }
} catch (System.InvalidOperationException outer) {
    Console.WriteLine(outer.Message);
    Console.WriteLine(outer.InnerException.Message);
}
"#
        ),
        &["wrapped", "root cause"]
    );
}
