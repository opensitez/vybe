crate::js_cases! {
    error_without_message_has_default_name_and_empty_message => {
        r#"
const e = new Error();
console.log(e.name);
console.log(e.message);
"#,
        ["Error", ""]
    };

    error_with_message_preserves_message => {
        r#"
const e = new Error("boom");
console.log(e.name);
console.log(e.message);
"#,
        ["Error", "boom"]
    };

    typeerror_defaults_name_correctly => {
        r#"
const e = new TypeError("bad");
console.log(e.name);
console.log(e.message);
"#,
        ["TypeError", "bad"]
    };

    rangeerror_defaults_name_correctly => {
        r#"
const e = new RangeError("out");
console.log(e.name);
console.log(e.message);
"#,
        ["RangeError", "out"]
    };

    referenceerror_defaults_name_correctly => {
        r#"
const e = new ReferenceError("missing");
console.log(e.name);
console.log(e.message);
"#,
        ["ReferenceError", "missing"]
    };

    syntaxerror_defaults_name_correctly => {
        r#"
const e = new SyntaxError("bad syntax");
console.log(e.name);
console.log(e.message);
"#,
        ["SyntaxError", "bad syntax"]
    };

    error_constructor_without_new_returns_error_object => {
        r#"
const e = Error("boom");
console.log(e instanceof Error);
console.log(e.message);
"#,
        ["true", "boom"]
    };

    error_is_instanceof_error => {
        r#"
console.log(new Error("x") instanceof Error);
"#,
        ["true"]
    };

    typeerror_is_instanceof_error => {
        r#"
console.log(new TypeError("x") instanceof Error);
"#,
        ["true"]
    };

    rangeerror_is_instanceof_error => {
        r#"
console.log(new RangeError("x") instanceof Error);
"#,
        ["true"]
    };

    error_tostring_without_message_uses_name_only => {
        r#"
console.log(new Error().toString());
"#,
        ["Error"]
    };

    error_tostring_with_message_uses_name_and_message => {
        r#"
console.log(new Error("boom").toString());
"#,
        ["Error: boom"]
    };

    error_tostring_with_custom_name_uses_custom_name => {
        r#"
const e = new Error("boom");
e.name = "CustomError";
console.log(e.toString());
"#,
        ["CustomError: boom"]
    };

    error_tostring_with_empty_message_uses_name_only => {
        r#"
const e = new Error("");
console.log(e.toString());
"#,
        ["Error"]
    };

    error_cause_object_is_preserved => {
        r#"
const inner = new Error("inner");
const outer = new Error("outer", { cause: inner });
console.log(outer.cause === inner);
"#,
        ["true"]
    };

    error_cause_primitive_is_preserved => {
        r#"
const e = new Error("boom", { cause: 42 });
console.log(e.cause);
"#,
        ["42"]
    };

    aggregateerror_exposes_errors_array => {
        r#"
const e = new AggregateError([1, 2, 3], "many");
console.log(e.errors.length);
console.log(e.message);
"#,
        ["3", "many"]
    };

    aggregateerror_is_instanceof_error => {
        r#"
console.log(new AggregateError([], "x") instanceof Error);
"#,
        ["true"]
    };

    error_message_coerces_number_to_string => {
        r#"
const e = new Error(42);
console.log(e.message);
"#,
        ["42"]
    };

    object_prototype_tostring_on_error_reports_error_tag => {
        r#"
console.log(Object.prototype.toString.call(new Error("x")));
"#,
        ["[object Error]"]
    };

    thrown_error_roundtrips_through_catch => {
        r#"
try {
  throw new Error("boom");
} catch (e) {
  console.log(e.message);
}
"#,
        ["boom"]
    };

    custom_error_name_is_writable => {
        r#"
const e = new Error("boom");
e.name = "BoomError";
console.log(e.name);
"#,
        ["BoomError"]
    };

    custom_error_message_is_writable => {
        r#"
const e = new Error("boom");
e.message = "changed";
console.log(e.message);
"#,
        ["changed"]
    };

    urierror_defaults_name_correctly => {
        r#"
const e = new URIError("bad uri");
console.log(e.name);
"#,
        ["URIError"]
    };

    evalerror_defaults_name_correctly => {
        r#"
const e = new EvalError("bad eval");
console.log(e.name);
"#,
        ["EvalError"]
    };

    error_tostring_with_empty_name_uses_message_only => {
        r#"
const e = new Error("boom");
e.name = "";
console.log(e.toString());
"#,
        ["boom"]
    };
}
