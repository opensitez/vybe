use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: `switch` Statement Matching, Fallthrough & Default Execution
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_switch_strict_equality_matching() {
    let src = r#"
const val = "5";
let res = "";
switch(val) {
    case 5: res = "number"; break;
    case "5": res = "string"; break;
}
console.log(res);
"#;
    assert_eq!(run_js(src), vec!["string"]); // Uses strict equality (===)!
}

#[test]
fn test_js_switch_case_fallthrough_without_break() {
    let src = r#"
const log = [];
switch(1) {
    case 1: log.push("c1");
    case 2: log.push("c2");
    case 3: log.push("c3"); break;
    case 4: log.push("c4");
}
console.log(log.join(","));
"#;
    assert_eq!(run_js(src), vec!["c1,c2,c3"]);
}

#[test]
fn test_js_switch_default_case_execution() {
    let src = r#"
let res = "";
switch(99) {
    case 1: res = "one"; break;
    default: res = "defaultVal"; break;
}
console.log(res);
"#;
    assert_eq!(run_js(src), vec!["defaultVal"]);
}

#[test]
fn test_js_switch_default_case_placement_in_middle_with_fallthrough() {
    let src = r#"
const log = [];
switch(99) {
    case 1: log.push("c1"); break;
    default: log.push("def"); // Default matches, falls through to case 2!
    case 2: log.push("c2"); break;
}
console.log(log.join(","));
"#;
    assert_eq!(run_js(src), vec!["def,c2"]);
}

#[test]
fn test_js_switch_expression_evaluated_once() {
    let src = r#"
let evalCount = 0;
const getVal = () => { evalCount++; return 2; };
switch(getVal()) {
    case 1: break;
    case 2: break;
}
console.log(evalCount);
"#;
    assert_eq!(run_js(src), vec!["1"]);
}

#[test]
fn test_js_switch_case_expressions_evaluated_lazily() {
    let src = r#"
const log = [];
const getVal = (n) => { log.push(`case${n}`); return n; };
switch(1) {
    case getVal(1): log.push("matched1"); break;
    case getVal(2): log.push("matched2"); break; // Case 2 is NOT evaluated because Case 1 matched and broke!
}
console.log(log.join(","));
"#;
    assert_eq!(run_js(src), vec!["case1,matched1"]);
}

#[test]
fn test_js_switch_matching_object_reference() {
    let src = r#"
const obj1 = { id: 1 };
const obj2 = { id: 1 };
let matched = false;
switch(obj1) {
    case obj2: matched = false; break;
    case obj1: matched = true; break;
}
console.log(matched);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_switch_matching_nan_is_false() {
    let src = r#"
let matched = false;
switch(NaN) {
    case NaN: matched = true; break; // NaN === NaN is false in switch!
    default: matched = false; break;
}
console.log(matched);
"#;
    assert_eq!(run_js(src), vec!["false"]);
}

#[test]
fn test_js_switch_matching_null_and_undefined() {
    let src = r#"
let res = "";
switch(null) {
    case undefined: res = "undef"; break;
    case null: res = "nullVal"; break;
}
console.log(res);
"#;
    assert_eq!(run_js(src), vec!["nullVal"]);
}

#[test]
fn test_js_switch_matching_boolean_true_range_pattern() {
    let src = r#"
const score = 85;
let grade = "";
switch(true) {
    case score >= 90: grade = "A"; break;
    case score >= 80: grade = "B"; break;
    default: grade = "F"; break;
}
console.log(grade);
"#;
    assert_eq!(run_js(src), vec!["B"]);
}

#[test]
fn test_js_switch_multiple_default_cases_throws_syntaxerror() {
    let src = r#"
try {
    eval("switch(1) { default: break; default: break; }");
} catch (e) {
    console.log("Multiple Defaults SyntaxError");
}
"#;
    assert_eq!(run_js(src), vec!["Multiple Defaults SyntaxError"]);
}

#[test]
fn test_js_switch_return_statement_inside_function() {
    let src = r#"
function fn(x) {
    switch(x) {
        case 1: return "one";
        case 2: return "two";
        default: return "other";
    }
}
console.log(fn(2));
"#;
    assert_eq!(run_js(src), vec!["two"]);
}

#[test]
fn test_js_switch_continue_statement_inside_loop() {
    let src = r#"
const log = [];
for (let i = 1; i <= 3; i++) {
    switch(i) {
        case 2: continue; // 'continue' inside switch targets enclosing loop!
    }
    log.push(i);
}
console.log(log.join(","));
"#;
    assert_eq!(run_js(src), vec!["1,3"]);
}

#[test]
fn test_js_switch_symbol_matching() {
    let src = r#"
const s1 = Symbol("a");
const s2 = Symbol("b");
let res = "";
switch(s1) {
    case s2: res = "s2"; break;
    case s1: res = "s1"; break;
}
console.log(res);
"#;
    assert_eq!(run_js(src), vec!["s1"]);
}

#[test]
fn test_js_switch_bigint_matching() {
    let src = r#"
let res = "";
switch(10n) {
    case 10: res = "number"; break;
    case 10n: res = "bigint"; break;
}
console.log(res);
"#;
    assert_eq!(run_js(src), vec!["bigint"]);
}

#[test]
fn test_js_switch_empty_switch_statement() {
    let src = r#"
let executed = true;
switch(1) {}
console.log(executed);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_switch_case_without_statement() {
    let src = r#"
const log = [];
switch(1) {
    case 1:
    case 2:
        log.push("matched1or2");
        break;
}
console.log(log.join(","));
"#;
    assert_eq!(run_js(src), vec!["matched1or2"]);
}

#[test]
fn test_js_switch_completion_value_in_eval() {
    let src = r#"
console.log(eval("switch(1) { case 1: 'val1'; 'val2'; }"));
"#;
    assert_eq!(run_js(src), vec!["val2"]);
}

#[test]
fn test_js_switch_case_with_comma_operator_expression() {
    let src = r#"
let res = "";
switch(5) {
    case (1, 5): res = "matched5"; break;
}
console.log(res);
"#;
    assert_eq!(run_js(src), vec!["matched5"]);
}

#[test]
fn test_js_switch_case_expression_side_effects_before_match() {
    let src = r#"
let sideEffects = 0;
switch(2) {
    case (sideEffects++, 1): break;
    case (sideEffects++, 2): break;
    case (sideEffects++, 3): break;
}
console.log(sideEffects);
"#;
    assert_eq!(run_js(src), vec!["2"]);
}

#[test]
fn test_js_switch_duplicate_case_values_match_first_clause() {
    let src = r#"
const out = [];
switch (3) {
    case 1: out.push("first");
    case 3: out.push("firstMatch"); break;
    case 3: out.push("secondMatch"); break;
    default: out.push("default");
}
console.log(out.join("|"));
"#;
    assert_eq!(run_js(src), vec!["firstMatch"]);
}

#[test]
fn test_js_switch_case_block_scoped_binding_visibility() {
    let src = r#"
const events = [];
switch (1) {
    case 1: {
        const local = "inside";
        events.push(local);
        break;
    }
    default:
        events.push("default");
}

let leaked = false;
try {
    local;
} catch (e) {
    leaked = e instanceof ReferenceError;
}

events.push(String(leaked));
console.log(events.join("|"));
"#;
    assert_eq!(run_js(src), vec!["inside|true"]);
}

#[test]
fn test_js_switch_fallthrough_with_block_scoped_bindings() {
    let src = r#"
const out = [];
switch("x") {
    case "x": {
        const marker = "A";
        out.push(marker);
    }
    case "y": {
        const marker = "B";
        out.push(marker);
        break;
    }
}
console.log(out.join("|"));
"#;
    assert_eq!(run_js(src), vec!["A|B"]);
}

#[test]
fn test_js_switch_case_fallthrough_crosses_default() {
    let src = r#"
const log = [];
switch (1) {
    case 1: log.push("c1");
    default: log.push("def");
    case 2: log.push("c2"); break;
}
console.log(log.join(","));
"#;
    assert_eq!(run_js(src), vec!["c1,def,c2"]);
}

#[test]
fn test_js_switch_break_inside_try_executes_finally() {
    let src = r#"
const log = [];
switch (1) {
    case 1:
        try {
            log.push("try");
            break;
        } finally {
            log.push("finally");
        }
    case 2:
        log.push("c2");
        break;
}
console.log(log.join(","));
"#;
    assert_eq!(run_js(src), vec!["try,finally"]);
}

#[test]
fn test_js_switch_zeros_match() {
    let src = r#"
let res = "";
switch(-0) {
    case +0: res = "matched"; break;
}
console.log(res);
"#;
    assert_eq!(run_js(src), vec!["matched"]);
}
