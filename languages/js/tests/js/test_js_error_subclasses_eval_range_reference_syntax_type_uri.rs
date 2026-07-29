use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Standard Builtin Error Subclasses (Eval, Range, Reference, Syntax, Type, URI)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_eval_error_construction() {
    let src = r#"
const err = new EvalError("Eval Failed");
console.log(err.name + "|" + err.message + "|" + (err instanceof Error));
"#;
    assert_eq!(run_js(src), vec!["EvalError|Eval Failed|true"]);
}

#[test]
fn test_js_range_error_construction() {
    let src = r#"
const err = new RangeError("Index Out of Bounds");
console.log(err.name + "|" + err.message + "|" + (err instanceof Error));
"#;
    assert_eq!(run_js(src), vec!["RangeError|Index Out of Bounds|true"]);
}

#[test]
fn test_js_reference_error_construction() {
    let src = r#"
const err = new ReferenceError("Variable Not Defined");
console.log(err.name + "|" + err.message + "|" + (err instanceof Error));
"#;
    assert_eq!(
        run_js(src),
        vec!["ReferenceError|Variable Not Defined|true"]
    );
}

#[test]
fn test_js_syntax_error_construction() {
    let src = r#"
const err = new SyntaxError("Unexpected Token");
console.log(err.name + "|" + err.message + "|" + (err instanceof Error));
"#;
    assert_eq!(run_js(src), vec!["SyntaxError|Unexpected Token|true"]);
}

#[test]
fn test_js_type_error_construction() {
    let src = r#"
const err = new TypeError("Invalid Type");
console.log(err.name + "|" + err.message + "|" + (err instanceof Error));
"#;
    assert_eq!(run_js(src), vec!["TypeError|Invalid Type|true"]);
}

#[test]
fn test_js_uri_error_construction() {
    let src = r#"
const err = new URIError("Malformed URI Sequence");
console.log(err.name + "|" + err.message + "|" + (err instanceof Error));
"#;
    assert_eq!(run_js(src), vec!["URIError|Malformed URI Sequence|true"]);
}

#[test]
fn test_js_error_subclasses_call_without_new() {
    let src = r#"
const e1 = RangeError("Range");
const e2 = TypeError("Type");
console.log((e1 instanceof RangeError) + "|" + (e2 instanceof TypeError));
"#;
    assert_eq!(run_js(src), vec!["true|true"]);
}

#[test]
fn test_js_range_error_array_length_negative_runtime_trigger() {
    let src = r#"
try {
    new Array(-10);
} catch (e) {
    console.log(e.name + "|" + (e instanceof RangeError));
}
"#;
    assert_eq!(run_js(src), vec!["RangeError|true"]);
}

#[test]
fn test_js_reference_error_undeclared_variable_runtime_trigger() {
    let src = r#"
try {
    nonExistentVariableReference;
} catch (e) {
    console.log(e.name + "|" + (e instanceof ReferenceError));
}
"#;
    assert_eq!(run_js(src), vec!["ReferenceError|true"]);
}

#[test]
fn test_js_type_error_null_property_access_runtime_trigger() {
    let src = r#"
try {
    null.foo();
} catch (e) {
    console.log(e.name + "|" + (e instanceof TypeError));
}
"#;
    assert_eq!(run_js(src), vec!["TypeError|true"]);
}

#[test]
fn test_js_uri_error_decode_uri_component_malformed_trigger() {
    let src = r#"
try {
    decodeURIComponent("%");
} catch (e) {
    console.log(e.name + "|" + (e instanceof URIError));
}
"#;
    assert_eq!(run_js(src), vec!["URIError|true"]);
}

#[test]
fn test_js_syntax_error_eval_invalid_code_trigger() {
    let src = r#"
try {
    eval("foo bar");
} catch (e) {
    console.log(e.name + "|" + (e instanceof SyntaxError));
}
"#;
    assert_eq!(run_js(src), vec!["SyntaxError|true"]);
}

#[test]
fn test_js_error_subclasses_prototype_name_properties() {
    let src = r#"
console.log([
    EvalError.prototype.name,
    RangeError.prototype.name,
    ReferenceError.prototype.name,
    SyntaxError.prototype.name,
    TypeError.prototype.name,
    URIError.prototype.name
].join(","));
"#;
    assert_eq!(
        run_js(src),
        vec!["EvalError,RangeError,ReferenceError,SyntaxError,TypeError,URIError"]
    );
}

#[test]
fn test_js_custom_error_subclass_extending_type_error() {
    let src = r#"
class ValidationError extends TypeError {
    constructor(field, message) {
        super(`${field}: ${message}`);
        this.name = "ValidationError";
        this.field = field;
    }
}
const err = new ValidationError("email", "Invalid format");
console.log(err.name + "|" + err.field + "|" + err.message + "|isTypeErr=" + (err instanceof TypeError));
"#;
    assert_eq!(
        run_js(src),
        vec!["ValidationError|email|email: Invalid format|isTypeErr=true"]
    );
}

#[test]
fn test_js_error_subclasses_tostring_method() {
    let src = r#"
const err = new RangeError("Number too large");
console.log(err.toString());
"#;
    assert_eq!(run_js(src), vec!["RangeError: Number too large"]);
}

#[test]
fn test_js_error_subclasses_message_empty_defaults() {
    let src = r#"
const err = new TypeError();
console.log(err.message === "");
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_error_subclasses_cause_support() {
    let src = r#"
const root = new Error("Disk Full");
const err = new URIError("Save Failed", { cause: root });
console.log(err.cause.message);
"#;
    assert_eq!(run_js(src), vec!["Disk Full"]);
}

#[test]
fn test_js_error_subclasses_prototype_chain_verification() {
    let src = r#"
const err = new RangeError("Range");
console.log(Object.getPrototypeOf(RangeError.prototype) === Error.prototype);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_error_subclasses_custom_name_override() {
    let src = r#"
const err = new SyntaxError("Bad Code");
err.name = "CustomSyntaxError";
console.log(err.toString());
"#;
    assert_eq!(run_js(src), vec!["CustomSyntaxError: Bad Code"]);
}

#[test]
fn test_js_type_error_assignment_to_const_runtime_trigger() {
    let src = r#"
const c = 10;
try {
    eval("c = 20;");
} catch (e) {
    console.log(e.name + "|isType=" + (e instanceof TypeError));
}
"#;
    assert_eq!(run_js(src), vec!["TypeError|isType=true"]);
}

#[test]
fn test_js_error_prototype_parent_is_object_prototype() {
    let src = r#"
console.log(Object.getPrototypeOf(Error.prototype) === Object.prototype);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

