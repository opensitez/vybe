/// Error.cause, AggregateError, custom error hierarchies, error re-wrapping,
/// SuppressedError (explicit resource management), stack-trace patterns.
use super::helpers::run_js;

// ── Error.cause ───────────────────────────────────────────────────────────────

#[test]
fn error_cause_option_is_accessible() {
    assert_eq!(
        run_js(
            r#"
const original = new TypeError("original issue");
const wrapped = new Error("high-level failure", { cause: original });
console.log(wrapped.message);
console.log(wrapped.cause instanceof TypeError);
console.log(wrapped.cause.message);
"#
        ),
        vec!["high-level failure", "true", "original issue"]
    );
}

#[test]
fn error_cause_can_be_any_value() {
    assert_eq!(
        run_js(
            r#"
const err = new Error("bad input", { cause: 42 });
console.log(err.cause);
"#
        ),
        vec!["42"]
    );
}

#[test]
fn error_cause_chain_three_levels() {
    assert_eq!(
        run_js(
            r#"
const root = new Error("root");
const mid = new Error("middle", { cause: root });
const top = new Error("top", { cause: mid });
console.log(top.cause.cause.message);
"#
        ),
        vec!["root"]
    );
}

#[test]
fn error_without_cause_has_undefined_cause() {
    assert_eq!(
        run_js(
            r#"
const err = new Error("simple");
console.log(err.cause);
"#
        ),
        vec!["undefined"]
    );
}

#[test]
fn error_cause_works_on_all_error_subtypes() {
    assert_eq!(
        run_js(
            r#"
const cause = new Error("root cause");
const types = [TypeError, RangeError, ReferenceError, SyntaxError, URIError, EvalError];
for (const T of types) {
    const e = new T("wrapper", { cause });
    console.log(e.cause === cause);
}
"#
        ),
        vec!["true", "true", "true", "true", "true", "true"]
    );
}

// ── AggregateError ────────────────────────────────────────────────────────────

#[test]
fn aggregate_error_holds_multiple_errors() {
    assert_eq!(
        run_js(
            r#"
const err = new AggregateError([
    new Error("first"),
    new Error("second"),
    new Error("third")
], "Multiple errors occurred");
console.log(err.message);
console.log(err.errors.length);
console.log(err.errors[0].message);
console.log(err.errors[2].message);
"#
        ),
        vec!["Multiple errors occurred", "3", "first", "third"]
    );
}

#[test]
fn aggregate_error_instanceof_error() {
    assert_eq!(
        run_js(
            r#"
const err = new AggregateError([], "test");
console.log(err instanceof AggregateError);
console.log(err instanceof Error);
console.log(err.name);
"#
        ),
        vec!["true", "true", "AggregateError"]
    );
}

#[test]
fn aggregate_error_with_cause() {
    assert_eq!(
        run_js(
            r#"
const root = new Error("root");
const agg = new AggregateError([new Error("e1")], "agg", { cause: root });
console.log(agg.cause.message);
"#
        ),
        vec!["root"]
    );
}

#[test]
fn promise_any_rejects_with_aggregate_error() {
    assert_eq!(
        run_js(
            r#"
async function main() {
    try {
        await Promise.any([
            Promise.reject(new Error("a")),
            Promise.reject(new Error("b"))
        ]);
    } catch (e) {
        console.log(e instanceof AggregateError);
        console.log(e.errors.length);
        console.log(e.errors[0].message);
    }
}
main();
"#
        ),
        vec!["true", "2", "a"]
    );
}

// ── custom error hierarchy ────────────────────────────────────────────────────

#[test]
fn custom_error_class_extends_error() {
    assert_eq!(
        run_js(
            r#"
class ValidationError extends Error {
    constructor(message, field) {
        super(message);
        this.name = "ValidationError";
        this.field = field;
    }
}
const e = new ValidationError("required", "email");
console.log(e instanceof ValidationError);
console.log(e instanceof Error);
console.log(e.name);
console.log(e.field);
console.log(e.message);
"#
        ),
        vec!["true", "true", "ValidationError", "email", "required"]
    );
}

#[test]
fn custom_error_hierarchy_three_levels() {
    assert_eq!(
        run_js(
            r#"
class AppError extends Error { constructor(msg) { super(msg); this.name = "AppError"; } }
class NetworkError extends AppError { constructor(msg, code) { super(msg); this.name = "NetworkError"; this.code = code; } }
class TimeoutError extends NetworkError { constructor() { super("request timed out", 408); this.name = "TimeoutError"; } }
const e = new TimeoutError();
console.log(e instanceof TimeoutError);
console.log(e instanceof NetworkError);
console.log(e instanceof AppError);
console.log(e instanceof Error);
console.log(e.code);
"#
        ),
        vec!["true", "true", "true", "true", "408"]
    );
}

#[test]
fn custom_error_with_cause_in_hierarchy() {
    assert_eq!(
        run_js(
            r#"
class ServiceError extends Error {
    constructor(message, options) {
        super(message, options);
        this.name = "ServiceError";
    }
}
const root = new Error("DB connection failed");
const svc = new ServiceError("Cannot fetch user", { cause: root });
console.log(svc.name);
console.log(svc.cause.message);
"#
        ),
        vec!["ServiceError", "DB connection failed"]
    );
}

// ── error re-wrapping patterns ────────────────────────────────────────────────

#[test]
fn wrapping_unknown_errors_in_known_type() {
    assert_eq!(
        run_js(
            r#"
function safeDivide(a, b) {
    if (b === 0) throw new RangeError("division by zero");
    return a / b;
}
function compute(a, b) {
    try { return safeDivide(a, b); }
    catch (e) { throw new Error("compute failed", { cause: e }); }
}
let msg = "";
try { compute(10, 0); }
catch (e) { msg = e.message + ":" + e.cause.constructor.name; }
console.log(msg);
"#
        ),
        vec!["compute failed:RangeError"]
    );
}

// ── error.name is configurable ────────────────────────────────────────────────

#[test]
fn error_name_can_be_customized_on_instance() {
    assert_eq!(
        run_js(
            r#"
const e = new Error("test");
e.name = "CustomError";
console.log(e.name);
console.log(e.message);
"#
        ),
        vec!["CustomError", "test"]
    );
}

// ── catching by type ──────────────────────────────────────────────────────────

#[test]
fn catch_and_rethrow_unknown_error() {
    assert_eq!(
        run_js(
            r#"
class DatabaseError extends Error {}
let result = "";
try {
    try { throw new TypeError("unexpected type"); }
    catch (e) {
        if (e instanceof DatabaseError) result = "db";
        else throw e;
    }
} catch (e) {
    result = "rethrown:" + e.constructor.name;
}
console.log(result);
"#
        ),
        vec!["rethrown:TypeError"]
    );
}

#[test]
fn aggregate_error_call_without_new_returns_instance() {
    assert_eq!(
        run_js(
            r#"
const e = AggregateError([new Error("a")], "msg");
console.log(e instanceof AggregateError);
console.log(e.message);
"#
        ),
        vec!["true", "msg"]
    );
}
