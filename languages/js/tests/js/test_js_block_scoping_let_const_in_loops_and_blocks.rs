use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Block Scoping (`let`, `const`) in Loops & Nested Blocks
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_block_scoped_let_in_if_statement() {
    let src = r#"
if (true) {
    let blockVar = "inside";
    console.log(blockVar);
}
console.log(typeof blockVar);
"#;
    assert_eq!(run_js(src), vec!["inside", "undefined"]);
}

#[test]
fn test_js_block_scoped_const_in_while_loop() {
    let src = r#"
let i = 0;
while (i < 2) {
    const loopConst = i * 10;
    console.log(loopConst);
    i++;
}
"#;
    assert_eq!(run_js(src), vec!["0", "10"]);
}

#[test]
fn test_js_for_loop_let_per_iteration_environment() {
    let src = r#"
const funcs = [];
for (let i = 0; i < 3; i++) {
    funcs.push(() => i);
}
console.log(funcs.map(f => f()).join(","));
"#;
    assert_eq!(run_js(src), vec!["0,1,2"]);
}

#[test]
fn test_js_for_in_loop_const_binding() {
    let src = r#"
const obj = { a: 1, b: 2 };
const keys = [];
for (const k in obj) {
    keys.push(k);
}
console.log(keys.join(","));
"#;
    assert_eq!(run_js(src), vec!["a,b"]);
}

#[test]
fn test_js_for_of_loop_const_binding() {
    let src = r#"
const arr = [10, 20];
const vals = [];
for (const x of arr) {
    vals.push(x * 2);
}
console.log(vals.join(","));
"#;
    assert_eq!(run_js(src), vec!["20,40"]);
}

#[test]
fn test_js_nested_block_shadowing() {
    let src = r#"
const x = "outer";
{
    const x = "middle";
    {
        const x = "inner";
        console.log(x);
    }
    console.log(x);
}
console.log(x);
"#;
    assert_eq!(run_js(src), vec!["inner", "middle", "outer"]);
}

#[test]
fn test_js_var_function_scope_ignores_block_boundaries() {
    let src = r#"
{
    var funcVar = "leaked";
}
console.log(funcVar);
"#;
    assert_eq!(run_js(src), vec!["leaked"]);
}

#[test]
fn test_js_switch_statement_shared_lexical_block_scope() {
    let src = r#"
try {
    eval("switch(1) { case 1: let a = 1; break; case 2: let a = 2; break; }"); // Redeclaring 'a' across cases without block braces is SyntaxError!
} catch (e) {
    console.log("Switch Redeclare SyntaxError");
}
"#;
    assert_eq!(run_js(src), vec!["Switch Redeclare SyntaxError"]);
}

#[test]
fn test_js_switch_case_isolated_block_scopes() {
    let src = r#"
switch(1) {
    case 1: {
        let a = 1;
        console.log(a);
        break;
    }
    case 2: {
        let a = 2;
        console.log(a);
        break;
    }
}
"#;
    assert_eq!(run_js(src), vec!["1"]);
}

#[test]
fn test_js_try_catch_block_scoping() {
    let src = r#"
try {
    throw new Error("Err");
} catch (e) {
    let catchLet = "insideCatch";
    console.log(catchLet);
}
console.log(typeof catchLet);
"#;
    assert_eq!(run_js(src), vec!["insideCatch", "undefined"]);
}

#[test]
fn test_js_for_loop_head_let_shadows_outer_variable() {
    let src = r#"
const i = "outerI";
for (let i = 0; i < 1; i++) {
    console.log(i);
}
console.log(i);
"#;
    assert_eq!(run_js(src), vec!["0", "outerI"]);
}

#[test]
fn test_js_for_loop_body_scope_shadows_loop_head_variable() {
    let src = r#"
const log = [];
for (let i = 0; i < 2; i++) {
    let i = "innerBody"; // Body scope shadows loop head 'i'!
    log.push(i);
}
console.log(log.join(","));
"#;
    assert_eq!(run_js(src), vec!["innerBody,innerBody"]);
}

#[test]
fn test_js_const_in_for_loop_head_throws_typeerror_on_reassignment() {
    let src = r#"
try {
    eval("for (const i = 0; i < 2; i++) {}"); // i++ attempts to reassign const in loop head!
} catch (e) {
    console.log("For Loop Const TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["For Loop Const TypeError"]);
}

#[test]
fn test_js_block_scoping_in_labelled_blocks() {
    let src = r#"
lbl: {
    const hidden = 100;
    break lbl;
}
console.log(typeof hidden);
"#;
    assert_eq!(run_js(src), vec!["undefined"]);
}

#[test]
fn test_js_standalone_block_statements() {
    let src = r#"
{
    const standalone = "valid";
    console.log(standalone);
}
console.log(typeof standalone);
"#;
    assert_eq!(run_js(src), vec!["valid", "undefined"]);
}

#[test]
fn test_js_function_parameter_scope_parent_of_body_scope() {
    let src = r#"
function fn(param = "defaultParam") {
    let paramVar = "bodyVar";
    console.log(`${param}:${paramVar}`);
}
fn();
"#;
    assert_eq!(run_js(src), vec!["defaultParam:bodyVar"]);
}

#[test]
fn test_js_function_parameter_shadowing_outer_scope() {
    let src = r#"
const val = "outer";
function fn(val) {
    console.log(val);
}
fn("innerParam");
"#;
    assert_eq!(run_js(src), vec!["innerParam"]);
}

#[test]
fn test_js_dead_code_elimination_block_scoping() {
    let src = r#"
if (false) {
    let unreachedLet = 1;
}
console.log(typeof unreachedLet);
"#;
    assert_eq!(run_js(src), vec!["undefined"]);
}

#[test]
fn test_js_dead_code_elimination_var_hoisting() {
    let src = r#"
if (false) {
    var unreachedVar = 1;
}
console.log(unreachedVar); // var declaration is hoisted even inside unreached if (false) block!
"#;
    assert_eq!(run_js(src), vec!["undefined"]);
}

#[test]
fn test_js_let_tdz_in_same_scope_before_initialization_throws() {
    let src = r#"
let hit = "no";
try {
    {
        console.log(x);
        let x = "too late";
    }
} catch (e) {
    hit = "error";
}
console.log(hit);
"#;
    assert_eq!(run_js(src), vec!["error"]);
}

#[test]
fn test_js_global_let_does_not_create_window_globalthis_property() {
    let src = r#"
let globalLetVar = "myLet";
console.log("globalLetVar" in globalThis); // Top-level let does NOT create property on globalThis!
"#;
    assert_eq!(run_js(src), vec!["false"]);
}

#[test]
fn test_js_global_var_creates_globalthis_property() {
    let src = r#"
var globalVarProp = "myVar";
console.log(globalThis.globalVarProp === "myVar");
"#;
    assert_eq!(run_js(src), vec!["true"]);
}
