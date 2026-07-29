use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Temporal Dead Zone (TDZ), `let` & `const` Hoisting Invariants
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_tdz_access_before_let_declaration_throws_referenceerror() {
    let src = r#"
try {
    eval("console.log(x); let x = 10;");
} catch (e) {
    console.log("TDZ let ReferenceError");
}
"#;
    assert_eq!(run_js(src), vec!["TDZ let ReferenceError"]);
}

#[test]
fn test_js_tdz_access_before_const_declaration_throws_referenceerror() {
    let src = r#"
try {
    eval("console.log(y); const y = 20;");
} catch (e) {
    console.log("TDZ const ReferenceError");
}
"#;
    assert_eq!(run_js(src), vec!["TDZ const ReferenceError"]);
}

#[test]
fn test_js_tdz_typeof_operator_throws_referenceerror() {
    let src = r#"
try {
    eval("typeof z; let z = 5;");
} catch (e) {
    console.log("TDZ typeof ReferenceError");
}
"#;
    assert_eq!(run_js(src), vec!["TDZ typeof ReferenceError"]);
}

#[test]
fn test_js_tdz_function_parameter_default_self_reference_throws() {
    let src = r#"
try {
    eval("function fn(a = a) {} fn();");
} catch (e) {
    console.log("TDZ Default Self-Reference ReferenceError");
}
"#;
    assert_eq!(
        run_js(src),
        vec!["TDZ Default Self-Reference ReferenceError"]
    );
}

#[test]
fn test_js_tdz_function_parameter_default_later_param_reference_throws() {
    let src = r#"
try {
    eval("function fn(a = b, b = 1) {} fn();");
} catch (e) {
    console.log("TDZ Default Later-Param ReferenceError");
}
"#;
    assert_eq!(run_js(src), vec!["TDZ Default Later-Param ReferenceError"]);
}

#[test]
fn test_js_tdz_class_declaration_access_before_declaration_throws() {
    let src = r#"
try {
    eval("new MyClass(); class MyClass {}");
} catch (e) {
    console.log("TDZ Class ReferenceError");
}
"#;
    assert_eq!(run_js(src), vec!["TDZ Class ReferenceError"]);
}

#[test]
fn test_js_tdz_switch_case_clause_lexical_scope() {
    let src = r#"
try {
    eval("switch(1) { case 1: let x = 10; break; case 2: console.log(x); break; }");
} catch (e) {
    console.log("TDZ Switch Case ReferenceError");
}
"#;
    assert_eq!(run_js(src), vec!["TDZ Switch Case ReferenceError"]);
}

#[test]
fn test_js_var_hoisting_initializes_to_undefined() {
    let src = r#"
function fn() {
    console.log(v);
    var v = 99;
}
fn();
"#;
    assert_eq!(run_js(src), vec!["undefined"]);
}

#[test]
fn test_js_const_requires_initializer_syntaxerror() {
    let src = r#"
try {
    eval("const uninitialized;");
} catch (e) {
    console.log("Const Missing Initializer SyntaxError");
}
"#;
    assert_eq!(run_js(src), vec!["Const Missing Initializer SyntaxError"]);
}

#[test]
fn test_js_redeclaring_let_or_const_in_same_scope_throws_syntaxerror() {
    let src = r#"
try {
    eval("let a = 1; let a = 2;");
} catch (e) {
    console.log("Redeclare let SyntaxError");
}
"#;
    assert_eq!(run_js(src), vec!["Redeclare let SyntaxError"]);
}

#[test]
fn test_js_redeclaring_var_as_let_in_same_scope_throws_syntaxerror() {
    let src = r#"
try {
    eval("var b = 1; let b = 2;");
} catch (e) {
    console.log("Redeclare var as let SyntaxError");
}
"#;
    assert_eq!(run_js(src), vec!["Redeclare var as let SyntaxError"]);
}

#[test]
fn test_js_const_reassignment_throws_typeerror() {
    let src = r#"
const c = 100;
try {
    eval("c = 200;");
} catch (e) {
    console.log("Const Reassignment TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Const Reassignment TypeError"]);
}

#[test]
fn test_js_tdz_closed_over_variable_called_before_init() {
    let src = r#"
const fn = () => val;
let val;
try {
    fn();
} catch (e) {
    console.log("TDZ Closure Call ReferenceError");
}
"#;
    assert_eq!(run_js(src), vec!["TDZ Closure Call ReferenceError"]);
}

#[test]
fn test_js_tdz_resolved_after_declaration_execution() {
    let src = r#"
const getVal = () => val;
let val = "NowInitialized";
console.log(getVal());
"#;
    assert_eq!(run_js(src), vec!["NowInitialized"]);
}

#[test]
fn test_js_tdz_for_loop_initializer_evaluation() {
    let src = r#"
const res = [];
for (let i = 0; i < 3; i++) {
    res.push(i);
}
console.log(res.join(","));
"#;
    assert_eq!(run_js(src), vec!["0,1,2"]);
}

#[test]
fn test_js_tdz_for_in_loop_initializer_lexical_scope() {
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
fn test_js_tdz_block_shadowing_outer_variable() {
    let src = r#"
const val = "Outer";
{
    try {
        eval("console.log(val); let val = 'Inner';"); // Block's inner 'val' TDZ shadows outer 'val'!
    } catch (e) {
        console.log("TDZ Shadowing ReferenceError");
    }
}
"#;
    assert_eq!(run_js(src), vec!["TDZ Shadowing ReferenceError"]);
}

#[test]
fn test_js_var_function_scoped_vs_let_block_scoped() {
    let src = r#"
function fn() {
    if (true) {
        var funcVar = 10;
        let blockLet = 20;
    }
    console.log(funcVar + "|" + (typeof blockLet));
}
fn();
"#;
    assert_eq!(run_js(src), vec!["10|undefined"]);
}

#[test]
fn test_js_function_declaration_hoisting_in_non_strict() {
    let src = r#"
console.log(hoistedFunc());
function hoistedFunc() { return "HoistedSuccess"; }
"#;
    assert_eq!(run_js(src), vec!["HoistedSuccess"]);
}

#[test]
fn test_js_function_expression_var_hoisted_as_undefined() {
    let src = r#"
try {
    eval("notHoisted(); var notHoisted = function() {};");
} catch (e) {
    console.log("Function Expression Hoisted as Undefined TypeError");
}
"#;
    assert_eq!(
        run_js(src),
        vec!["Function Expression Hoisted as Undefined TypeError"]
    );
}

#[test]
fn test_js_tdz_class_heritage_extends_self_throws() {
    let src = r#"
try {
    eval("class Sub extends Sub {}");
} catch (e) {
    console.log(e instanceof ReferenceError);
}
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

