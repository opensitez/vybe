use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: `AggregateError` & ES2022 `Error.cause` Chaining
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_aggregate_error_construction_and_errors_array() {
    let src = r#"
const err = new AggregateError([new Error("Err1"), new Error("Err2")], "Bulk Failure");
console.log(err.name + "|" + err.message + "|" + err.errors.length);
"#;
    assert_eq!(run_js(src), vec!["AggregateError|Bulk Failure|2"]);
}

#[test]
fn test_js_error_cause_option_property_es2022() {
    let src = r#"
const causeErr = new Error("LowLevelIOError");
const mainErr = new Error("Failed to process file", { cause: causeErr });
console.log(mainErr.cause.message);
"#;
    assert_eq!(run_js(src), vec!["LowLevelIOError"]);
}

#[test]
fn test_js_error_cause_primitive_value() {
    let src = r#"
const err = new TypeError("Invalid Config", { cause: 404 });
console.log(err.cause);
"#;
    assert_eq!(run_js(src), vec!["404"]);
}

#[test]
fn test_js_aggregate_error_cause_combination() {
    let src = r#"
const aggErr = new AggregateError([1, 2], "Operation failed", { cause: "RootCause" });
console.log(aggErr.message + "|cause=" + aggErr.cause + "|errors=" + aggErr.errors.join(","));
"#;
    assert_eq!(
        run_js(src),
        vec!["Operation failed|cause=RootCause|errors=1,2"]
    );
}

#[test]
fn test_js_error_cause_omitted_when_options_not_provided() {
    let src = r#"
const err = new Error("Regular");
console.log(Object.hasOwn(err, "cause"));
"#;
    assert_eq!(run_js(src), vec!["false"]);
}

#[test]
fn test_js_aggregate_error_iterable_errors_argument() {
    let src = r#"
const set = new Set(["ErrA", "ErrB"]);
const agg = new AggregateError(set, "SetErrors");
console.log(agg.errors.join(","));
"#;
    assert_eq!(run_js(src), vec!["ErrA,ErrB"]);
}

#[test]
fn test_js_aggregate_error_non_iterable_errors_throws_typeerror() {
    let src = r#"
try {
    new AggregateError(12345, "Bad");
} catch (e) {
    console.log("AggregateError Non-Iterable TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["AggregateError Non-Iterable TypeError"]);
}

#[test]
fn test_js_aggregate_error_prototype_name() {
    let src = r#"
console.log(AggregateError.prototype.name);
"#;
    assert_eq!(run_js(src), vec!["AggregateError"]);
}

#[test]
fn test_js_aggregate_error_instanceof_error() {
    let src = r#"
const agg = new AggregateError([], "Empty");
console.log((agg instanceof AggregateError) + "|" + (agg instanceof Error));
"#;
    assert_eq!(run_js(src), vec!["true|true"]);
}

#[test]
fn test_js_error_cause_object_identity() {
    let src = r#"
const original = { code: "ECONNRESET" };
const err = new Error("NetErr", { cause: original });
console.log(err.cause === original);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_error_cause_nested_chaining() {
    let src = r#"
const e1 = new Error("Level 1");
const e2 = new Error("Level 2", { cause: e1 });
const e3 = new Error("Level 3", { cause: e2 });
console.log(`${e3.message} -> ${e3.cause.message} -> ${e3.cause.cause.message}`);
"#;
    assert_eq!(run_js(src), vec!["Level 3 -> Level 2 -> Level 1"]);
}

#[test]
fn test_js_aggregate_error_promise_any_rejection() {
    let src = r#"
(async () => {
    try {
        await Promise.any([Promise.reject("FailA"), Promise.reject("FailB")]);
    } catch (e) {
        console.log((e instanceof AggregateError) + "|" + e.errors.join(","));
    }
})();
"#;
    assert_eq!(run_js(src), vec!["true|FailA,FailB"]);
}

#[test]
fn test_js_error_options_undefined_cause() {
    let src = r#"
const err = new Error("Msg", { cause: undefined });
console.log(Object.hasOwn(err, "cause") + "|" + (err.cause === undefined));
"#;
    assert_eq!(run_js(src), vec!["true|true"]);
}

#[test]
fn test_js_error_subclasses_cause_support() {
    let src = r#"
const cause = "BadType";
const err = new TypeError("Invalid arg", { cause });
console.log(err.cause);
"#;
    assert_eq!(run_js(src), vec!["BadType"]);
}

#[test]
fn test_js_aggregate_error_errors_array_is_copied() {
    let src = r#"
const input = [1, 2];
const agg = new AggregateError(input, "Msg");
input.push(3); // Modifying input array after construction
console.log(agg.errors.length); // agg.errors is frozen / copied snapshot
"#;
    assert_eq!(run_js(src), vec!["2"]);
}

#[test]
fn test_js_aggregate_error_empty_message_defaults_to_empty_string() {
    let src = r#"
const agg = new AggregateError([]);
console.log(agg.message === "");
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_error_tostring_formatting() {
    let src = r#"
const err = new Error("SampleMessage");
console.log(err.toString());
"#;
    assert_eq!(run_js(src), vec!["Error: SampleMessage"]);
}

#[test]
fn test_js_aggregate_error_tostring_formatting() {
    let src = r#"
const agg = new AggregateError([], "Bulk");
console.log(agg.toString());
"#;
    assert_eq!(run_js(src), vec!["AggregateError: Bulk"]);
}

#[test]
fn test_js_error_factory_call_without_new() {
    let src = r#"
const err = Error("NoNewKeyword");
console.log(err.message + "|" + (err instanceof Error));
"#;
    assert_eq!(run_js(src), vec!["NoNewKeyword|true"]);
}

#[test]
fn test_js_aggregate_error_factory_call_without_new() {
    let src = r#"
const agg = AggregateError([1], "NoNew");
console.log(agg.message + "|" + (agg instanceof AggregateError));
"#;
    assert_eq!(run_js(src), vec!["NoNew|true"]);
}
