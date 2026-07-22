use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Function Declaration Hoisting in Blocks & Strict vs Non-Strict Annex B
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_function_hoisting_top_level_function_statement() {
    let src = r#"
console.log(hoisted());
function hoisted() { return "TopLevelHoisted"; }
"#;
    assert_eq!(run_js(src), vec!["TopLevelHoisted"]);
}

#[test]
fn test_js_function_hoisting_inside_function_scope() {
    let src = r#"
function outer() {
    return inner();
    function inner() { return "InnerHoisted"; }
}
console.log(outer());
"#;
    assert_eq!(run_js(src), vec!["InnerHoisted"]);
}

#[test]
fn test_js_function_declaration_in_block_strict_mode() {
    let src = r#"
"use strict";
{
    function blockFunc() { return "BlockScopeStrict"; }
    console.log(blockFunc());
}
console.log(typeof blockFunc); // In strict mode, function declarations in blocks are strictly block-scoped!
"#;
    assert_eq!(run_js(src), vec!["BlockScopeStrict", "undefined"]);
}

#[test]
fn test_js_function_declaration_in_if_block_annex_b_non_strict() {
    let src = r#"
if (true) {
    function ifFunc() { return "IfBlockFunc"; }
}
console.log(ifFunc()); // In non-strict Annex B, function declaration in block hoists to enclosing function/global!
"#;
    assert_eq!(run_js(src), vec!["IfBlockFunc"]);
}

#[test]
fn test_js_function_declaration_in_false_if_block_annex_b_non_strict() {
    let src = r#"
console.log(typeof unexecutedFunc); // Var binding is hoisted (undefined), but assignment is skipped!
if (false) {
    function unexecutedFunc() {}
}
console.log(unexecutedFunc);
"#;
    assert_eq!(run_js(src), vec!["undefined", "undefined"]);
}

#[test]
fn test_js_function_declaration_overwrites_earlier_declaration() {
    let src = r#"
function fn() { return "first"; }
function fn() { return "second"; }
console.log(fn());
"#;
    assert_eq!(run_js(src), vec!["second"]);
}

#[test]
fn test_js_function_declaration_overwrites_var_declaration() {
    let src = r#"
var test = "varVal";
function test() { return "funcVal"; }
console.log(typeof test); // Function declaration is hoisted first, then var initialization 'test = "varVal"' assigns string!
"#;
    assert_eq!(run_js(src), vec!["string"]);
}

#[test]
fn test_js_var_declaration_does_not_overwrite_uninitialized_function() {
    let src = r#"
console.log(typeof test);
var test;
function test() {}
"#;
    assert_eq!(run_js(src), vec!["function"]);
}

#[test]
fn test_js_function_expression_var_hoisting_is_undefined() {
    let src = r#"
console.log(typeof expr);
var expr = function() { return "expr"; };
"#;
    assert_eq!(run_js(src), vec!["undefined"]);
}

#[test]
fn test_js_named_function_expression_name_only_in_own_scope() {
    let src = r#"
const fn = function internalName() {
    return typeof internalName;
};
console.log(fn() + "|" + (typeof internalName));
"#;
    assert_eq!(run_js(src), vec!["function|undefined"]);
}

#[test]
fn test_js_function_declaration_in_switch_case_block() {
    let src = r#"
switch(1) {
    case 1:
        function switchFunc() { return "SwitchFunc"; }
        break;
}
console.log(switchFunc());
"#;
    assert_eq!(run_js(src), vec!["SwitchFunc"]);
}

#[test]
fn test_js_function_declaration_in_while_loop_block_annex_b() {
    let src = r#"
let i = 0;
while (i < 1) {
    function whileFunc() { return "WhileFunc"; }
    i++;
}
console.log(whileFunc());
"#;
    assert_eq!(run_js(src), vec!["WhileFunc"]);
}

#[test]
fn test_js_function_declaration_redeclaring_let_throws_syntaxerror() {
    let src = r#"
try {
    eval("let a = 1; function a() {}");
} catch (e) {
    console.log("Redeclare let with Function SyntaxError");
}
"#;
    assert_eq!(run_js(src), vec!["Redeclare let with Function SyntaxError"]);
}

#[test]
fn test_js_function_declaration_redeclaring_const_throws_syntaxerror() {
    let src = r#"
try {
    eval("const c = 1; function c() {}");
} catch (e) {
    console.log("Redeclare const with Function SyntaxError");
}
"#;
    assert_eq!(
        run_js(src),
        vec!["Redeclare const with Function SyntaxError"]
    );
}

#[test]
fn test_js_function_declaration_parameter_shadowing() {
    let src = r#"
function fn(shadowed) {
    function shadowed() { return "InnerFunc"; }
    return shadowed();
}
console.log(fn("paramVal"));
"#;
    assert_eq!(run_js(src), vec!["InnerFunc"]);
}

#[test]
fn test_js_function_declaration_in_try_block_strict_mode() {
    let src = r#"
"use strict";
try {
    function tryFunc() { return "TryFunc"; }
    console.log(tryFunc());
} catch (e) {}
console.log(typeof tryFunc);
"#;
    assert_eq!(run_js(src), vec!["TryFunc", "undefined"]);
}

#[test]
fn test_js_function_declaration_in_catch_block_strict_mode() {
    let src = r#"
"use strict";
try {
    throw new Error();
} catch (e) {
    function catchFunc() { return "CatchFunc"; }
    console.log(catchFunc());
}
console.log(typeof catchFunc);
"#;
    assert_eq!(run_js(src), vec!["CatchFunc", "undefined"]);
}

#[test]
fn test_js_function_declaration_in_finally_block_strict_mode() {
    let src = r#"
"use strict";
try {} finally {
    function finallyFunc() { return "FinallyFunc"; }
    console.log(finallyFunc());
}
console.log(typeof finallyFunc);
"#;
    assert_eq!(run_js(src), vec!["FinallyFunc", "undefined"]);
}

#[test]
fn test_js_function_declaration_with_default_parameters_evaluation_order() {
    let src = r#"
function fn(a = getA()) {
    function getA() { return 100; } // getA is hoisted within body scope, NOT accessible to parameter default evaluation!
    return a;
}
try {
    fn();
} catch (e) {
    console.log("Parameter Default Function Hoisting ReferenceError");
}
"#;
    assert_eq!(
        run_js(src),
        vec!["Parameter Default Function Hoisting ReferenceError"]
    );
}

#[test]
fn test_js_function_declaration_in_eval_strict_mode() {
    let src = r#"
"use strict";
eval("function evalFunc() { return 'EvalStrict'; }");
console.log(typeof evalFunc); // In strict mode eval, functions declared inside eval stay in eval scope!
"#;
    assert_eq!(run_js(src), vec!["undefined"]);
}
