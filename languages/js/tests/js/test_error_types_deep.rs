/// Error types and custom errors — name, message, stack, cause, inheritance
use super::helpers::run_js;

#[test]
fn syntax_error_from_eval() {
    assert_eq!(
        run_js(
            r#"
let err = null;
try { eval("if ("); } catch (e) { err = e; }
console.log(err instanceof SyntaxError);
console.log(err instanceof Error);
"#
        ),
        vec!["true", "true"]
    );
}

#[test]
fn type_error_from_null_access() {
    assert_eq!(
        run_js(
            r#"
try { null.foo; } catch (e) {
    console.log(e instanceof TypeError);
    console.log(e.name);
}
"#
        ),
        vec!["true", "TypeError"]
    );
}

#[test]
fn range_error_from_invalid_array_length() {
    assert_eq!(
        run_js(
            r#"
try { new Array(-1); } catch (e) {
    console.log(e instanceof RangeError);
    console.log(e.name);
}
"#
        ),
        vec!["true", "RangeError"]
    );
}

#[test]
fn reference_error_from_undeclared() {
    assert_eq!(
        run_js(
            r#"
try { undeclaredVariable; } catch (e) {
    console.log(e instanceof ReferenceError);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn uri_error_from_decode() {
    assert_eq!(
        run_js(
            r#"
try { decodeURIComponent("%"); } catch (e) {
    console.log(e instanceof URIError);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn error_cause_property() {
    assert_eq!(
        run_js(
            r#"
const cause = new Error("root cause");
const err = new Error("outer", { cause });
console.log(err.message);
console.log(err.cause === cause);
console.log(err.cause.message);
"#
        ),
        vec!["outer", "true", "root cause"]
    );
}

#[test]
fn custom_error_class() {
    assert_eq!(
        run_js(
            r#"
class AppError extends Error {
    constructor(msg, code) {
        super(msg);
        this.name = "AppError";
        this.code = code;
    }
}
const e = new AppError("failed", 404);
console.log(e instanceof AppError);
console.log(e instanceof Error);
console.log(e.name);
console.log(e.code);
console.log(e.message);
"#
        ),
        vec!["true", "true", "AppError", "404", "failed"]
    );
}

#[test]
fn error_stack_is_string() {
    assert_eq!(
        run_js(
            r#"
const e = new Error("test");
console.log(typeof e.stack);
"#
        ),
        vec!["string"]
    );
}

#[test]
fn aggregate_error_errors_array() {
    assert_eq!(
        run_js(
            r#"
const agg = new AggregateError([new Error("a"), new Error("b")], "multiple errors");
console.log(agg.message);
console.log(agg.errors.length);
console.log(agg.errors[0].message);
"#
        ),
        vec!["multiple errors", "2", "a"]
    );
}

#[test]
fn error_to_string_format() {
    assert_eq!(
        run_js(
            r#"
const e = new Error("oops");
console.log(e.toString());
const noMsg = new Error("");
console.log(noMsg.toString());
"#
        ),
        vec!["Error: oops", "Error"]
    );
}

#[test]
fn catch_by_type_in_hierarchy() {
    assert_eq!(
        run_js(
            r#"
function risky(x) {
    if (x < 0) throw new RangeError("negative");
    if (typeof x !== "number") throw new TypeError("not a number");
}
function handle(x) {
    try {
        risky(x);
    } catch (e) {
        if (e instanceof RangeError) return "range";
        if (e instanceof TypeError) return "type";
        throw e;
    }
}
console.log(handle(-1));
console.log(handle("foo"));
"#
        ),
        vec!["range", "type"]
    );
}

#[test]
fn eval_error_exists() {
    assert_eq!(
        run_js(
            r#"
const e = new EvalError("test");
console.log(e instanceof EvalError);
console.log(e instanceof Error);
console.log(e.name);
"#
        ),
        vec!["true", "true", "EvalError"]
    );
}

#[test]
fn aggregate_error_is_array() {
    assert_eq!(
        run_js(
            r#"
const agg = new AggregateError([new Error("a")], "msg");
console.log(Array.isArray(agg.errors));
"#
        ),
        vec!["true"]
    );
}
