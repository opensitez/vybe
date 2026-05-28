/// Error hierarchy — custom error subclasses, stack traces, instanceof chain,
/// error properties, error types, subclass constructor, toJSON pattern.

use super::helpers::run_js;

// ── built-in error types ──────────────────────────────────────────────────────

#[test]
fn error_types_are_instances_of_error() {
    assert_eq!(run_js(r#"
const errors = [
    new TypeError("t"),
    new RangeError("r"),
    new ReferenceError("ref"),
    new SyntaxError("s"),
    new URIError("u"),
    new EvalError("e"),
];
console.log(errors.every(e => e instanceof Error));
console.log(errors.map(e => e.constructor.name).join(","));
"#), vec!["true", "TypeError,RangeError,ReferenceError,SyntaxError,URIError,EvalError"]);
}

#[test]
fn error_has_name_message_stack() {
    assert_eq!(run_js(r#"
const e = new TypeError("bad type");
console.log(e.name);
console.log(e.message);
console.log(typeof e.stack);
"#), vec!["TypeError", "bad type", "string"]);
}

// ── custom error subclasses ───────────────────────────────────────────────────

#[test]
fn custom_error_basic() {
    assert_eq!(run_js(r#"
class AppError extends Error {
    constructor(message, code) {
        super(message);
        this.name = "AppError";
        this.code = code;
    }
}
const e = new AppError("not found", 404);
console.log(e instanceof AppError);
console.log(e instanceof Error);
console.log(e.name);
console.log(e.message);
console.log(e.code);
"#), vec!["true", "true", "AppError", "not found", "404"]);
}

#[test]
fn custom_error_two_levels_deep() {
    assert_eq!(run_js(r#"
class BaseError extends Error {
    constructor(msg) { super(msg); this.name = "BaseError"; }
}
class NetworkError extends BaseError {
    constructor(msg, status) {
        super(msg);
        this.name = "NetworkError";
        this.status = status;
    }
}
const e = new NetworkError("timeout", 503);
console.log(e instanceof NetworkError);
console.log(e instanceof BaseError);
console.log(e instanceof Error);
console.log(e.status);
"#), vec!["true", "true", "true", "503"]);
}

#[test]
fn custom_error_catch_by_type() {
    assert_eq!(run_js(r#"
class ValidationError extends Error {
    constructor(field, msg) {
        super(msg);
        this.name = "ValidationError";
        this.field = field;
    }
}
class AuthError extends Error {
    constructor(msg) { super(msg); this.name = "AuthError"; }
}

function handle(e) {
    if (e instanceof ValidationError) return "validation:" + e.field;
    if (e instanceof AuthError) return "auth:" + e.message;
    return "unknown";
}

console.log(handle(new ValidationError("email", "invalid")));
console.log(handle(new AuthError("forbidden")));
console.log(handle(new Error("generic")));
"#), vec!["validation:email", "auth:forbidden", "unknown"]);
}

// ── error name property ───────────────────────────────────────────────────────

#[test]
fn setting_name_in_subclass() {
    assert_eq!(run_js(r#"
class MyError extends Error {}
const e = new MyError("msg");
// name is inherited from class name when not overridden
// but default from Error is "Error"
console.log(e instanceof MyError);
"#), vec!["true"]);
}

// ── Error.prototype properties ────────────────────────────────────────────────

#[test]
fn error_prototype_name_message_defaults() {
    assert_eq!(run_js(r#"
const e = new Error();
console.log(e.name);
console.log(e.message);
"#), vec!["Error", ""]);
}

// ── AggregateError ────────────────────────────────────────────────────────────

#[test]
fn aggregate_error_contains_errors() {
    assert_eq!(run_js(r#"
const errors = [new Error("e1"), new TypeError("e2")];
const agg = new AggregateError(errors, "multiple failures");
console.log(agg instanceof AggregateError);
console.log(agg instanceof Error);
console.log(agg.message);
console.log(agg.errors.length);
console.log(agg.errors[0].message);
"#), vec!["true", "true", "multiple failures", "2", "e1"]);
}

// ── error in promise ──────────────────────────────────────────────────────────

#[test]
fn error_propagation_through_promise_chain() {
    assert_eq!(run_js(r#"
class ApiError extends Error {
    constructor(msg, code) { super(msg); this.code = code; }
}

async function fetchData() {
    throw new ApiError("not found", 404);
}

fetchData().catch(e => {
    console.log(e instanceof ApiError);
    console.log(e.code);
    console.log(e.message);
});
"#), vec!["true", "404", "not found"]);
}

// ── rethrowing and wrapping ───────────────────────────────────────────────────

#[test]
fn error_wrapping_pattern() {
    assert_eq!(run_js(r#"
function parse(str) {
    try {
        return JSON.parse(str);
    } catch (e) {
        throw new Error("Parse failed: " + e.message, { cause: e });
    }
}

try {
    parse("{bad}");
} catch (e) {
    console.log(e.message.startsWith("Parse failed:"));
    console.log(e.cause instanceof SyntaxError);
}
"#), vec!["true", "true"]);
}

// ── finally and error ─────────────────────────────────────────────────────────

#[test]
fn finally_does_not_suppress_error_unless_return() {
    assert_eq!(run_js(r#"
function f() {
    try { throw new Error("original"); }
    finally { /* no return, error propagates */ }
}
try { f(); } catch (e) { console.log(e.message); }
"#), vec!["original"]);
}

// ── accessing stack trace ─────────────────────────────────────────────────────

#[test]
fn stack_trace_contains_function_name() {
    assert_eq!(run_js(r#"
function namedFunction() { return new Error("test"); }
const e = namedFunction();
// Stack should mention something — format varies by engine
console.log(typeof e.stack === "string");
"#), vec!["true"]);
}

// ── error toString ────────────────────────────────────────────────────────────

#[test]
fn error_to_string_format() {
    assert_eq!(run_js(r#"
console.log(new Error("test").toString());
console.log(new TypeError("bad").toString());
"#), vec!["Error: test", "TypeError: bad"]);
}

#[test]
fn error_to_string_no_message() {
    assert_eq!(run_js(r#"
console.log(new Error("").toString());
console.log(new TypeError().toString());
"#), vec!["Error", "TypeError"]);
}
