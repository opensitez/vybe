use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: `structuredClone` Error Serialization & Exception Clones
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_structured_clone_standard_error_object() {
    let src = r#"
const err = new Error("Sample Error Message");
const clone = structuredClone(err);
console.log((clone !== err) + "|" + (clone instanceof Error) + "|" + (clone.name === "Error") + "|" + (clone.message === "Sample Error Message"));
"#;
    assert_eq!(run_js(src), vec!["true|true|true|true"]);
}

#[test]
fn test_js_structured_clone_type_error_subclass() {
    let src = r#"
const err = new TypeError("Invalid argument");
const clone = structuredClone(err);
console.log((clone instanceof TypeError) + "|" + (clone.name === "TypeError") + "|" + (clone.message === "Invalid argument"));
"#;
    assert_eq!(run_js(src), vec!["true|true|true"]);
}

#[test]
fn test_js_structured_clone_range_error_subclass() {
    let src = r#"
const err = new RangeError("Out of bounds");
const clone = structuredClone(err);
console.log((clone instanceof RangeError) + "|" + clone.message);
"#;
    assert_eq!(run_js(src), vec!["true|Out of bounds"]);
}

#[test]
fn test_js_structured_clone_syntax_error_subclass() {
    let src = r#"
const err = new SyntaxError("Unexpected token");
const clone = structuredClone(err);
console.log((clone instanceof SyntaxError) + "|" + clone.message);
"#;
    assert_eq!(run_js(src), vec!["true|Unexpected token"]);
}

#[test]
fn test_js_structured_clone_eval_error_subclass() {
    let src = r#"
const err = new EvalError("Eval error");
const clone = structuredClone(err);
console.log((clone instanceof EvalError) + "|" + clone.message);
"#;
    assert_eq!(run_js(src), vec!["true|Eval error"]);
}

#[test]
fn test_js_structured_clone_reference_error_subclass() {
    let src = r#"
const err = new ReferenceError("Not defined");
const clone = structuredClone(err);
console.log((clone instanceof ReferenceError) + "|" + clone.message);
"#;
    assert_eq!(run_js(src), vec!["true|Not defined"]);
}

#[test]
fn test_js_structured_clone_uri_error_subclass() {
    let src = r#"
const err = new URIError("URI error");
const clone = structuredClone(err);
console.log((clone instanceof URIError) + "|" + clone.message);
"#;
    assert_eq!(run_js(src), vec!["true|URI error"]);
}

#[test]
fn test_js_structured_clone_aggregate_error() {
    let src = r#"
const err1 = new Error("Err1");
const err2 = new Error("Err2");
const agg = new AggregateError([err1, err2], "Bulk Failure");
const clone = structuredClone(agg);

console.log((clone instanceof AggregateError) + "|" + (clone.message === "Bulk Failure") + "|" + clone.errors.map(e => e.message).join(","));
"#;
    assert_eq!(run_js(src), vec!["true|true|Err1,Err2"]);
}

#[test]
fn test_js_structured_clone_error_cause_property_cloned() {
    let src = r#"
const causeErr = new Error("LowLevelCause");
const mainErr = new Error("HighLevelMain", { cause: causeErr });
const clone = structuredClone(mainErr);

console.log((clone.cause !== causeErr) + "|" + (clone.cause.message === "LowLevelCause"));
"#;
    assert_eq!(run_js(src), vec!["true|true"]);
}

#[test]
fn test_js_structured_clone_error_stack_property_cloned() {
    let src = r#"
const err = new Error("StackMsg");
const clone = structuredClone(err);
console.log(typeof clone.stack === "string" && clone.stack.includes("StackMsg"));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_structured_clone_custom_properties_on_error() {
    let src = r#"
const err = new Error("CustomError");
err.code = "ERR_NOT_FOUND";
err.status = 404;
const clone = structuredClone(err);

console.log(clone.code + "|" + clone.status);
"#;
    assert_eq!(run_js(src), vec!["ERR_NOT_FOUND|404"]);
}

#[test]
fn test_js_structured_clone_custom_error_class_prototype_fallback() {
    let src = r#"
class CustomAppError extends Error {
    constructor(msg) {
        super(msg);
        this.name = "CustomAppError";
    }
}
const err = new CustomAppError("AppFailed");
const clone = structuredClone(err);

console.log(clone.name + "|" + (clone instanceof Error) + "|isCustom=" + (clone instanceof CustomAppError));
"#;
    assert_eq!(run_js(src), vec!["CustomAppError|true|isCustom=false"]); // Structured clone serializes custom error as standard Error!
}

#[test]
fn test_js_structured_clone_suppressed_error() {
    let src = r#"
const p = new Error("Primary");
const s = new Error("Suppressed");
const err = new SuppressedError(p, s, "SuppressedMsg");
const clone = structuredClone(err);

console.log((clone instanceof SuppressedError) + "|" + clone.error.message + "|" + clone.suppressed.message);
"#;
    assert_eq!(run_js(src), vec!["true|Primary|Suppressed"]);
}

#[test]
fn test_js_structured_clone_error_with_cyclic_cause() {
    let src = r#"
const e1 = new Error("E1");
const e2 = new Error("E2", { cause: e1 });
e1.cause = e2;

const clone = structuredClone(e2);
console.log((clone.cause.cause === clone));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_structured_clone_error_with_non_serializable_cause_throws() {
    let src = r#"
const err = new Error("BadCauseMsg", { cause: () => {} });
try {
    structuredClone(err);
} catch (e) {
    console.log("DataCloneError Function Cause");
}
"#;
    assert_eq!(run_js(src), vec!["DataCloneError Function Cause"]);
}

#[test]
fn test_js_structured_clone_error_with_array_cause() {
    let src = r#"
const err = new Error("MultiCause", { cause: [10, 20] });
const clone = structuredClone(err);
console.log(Array.isArray(clone.cause) + "|" + clone.cause.join(","));
"#;
    assert_eq!(run_js(src), vec!["true|10,20"]);
}

#[test]
fn test_js_structured_clone_error_primitive_cause() {
    let src = r#"
const err = new Error("CodeError", { cause: 500 });
const clone = structuredClone(err);
console.log(clone.cause);
"#;
    assert_eq!(run_js(src), vec!["500"]);
}

#[test]
fn test_js_structured_clone_error_empty_message() {
    let src = r#"
const err = new Error();
const clone = structuredClone(err);
console.log(clone.message === "");
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_structured_clone_error_tostring_method() {
    let src = r#"
const err = new TypeError("Invalid Type");
const clone = structuredClone(err);
console.log(clone.toString());
"#;
    assert_eq!(run_js(src), vec!["TypeError: Invalid Type"]);
}

#[test]
fn test_js_structured_clone_error_inside_map_value() {
    let src = r#"
const errMap = new Map([["errKey", new Error("InMapError")]]);
const clone = structuredClone(errMap);
const clonedErr = clone.get("errKey");
console.log((clonedErr instanceof Error) + "|" + clonedErr.message);
"#;
    assert_eq!(run_js(src), vec!["true|InMapError"]);
}
