use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Direct vs. Indirect `eval()` Scope Binding Differences
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_direct_eval_local_variable_access() {
    let src = r#"
function fn() {
    const x = 42;
    return eval("x + 10");
}
console.log(fn());
"#;
    assert_eq!(run_js(src), vec!["52"]);
}

#[test]
fn test_js_indirect_eval_global_scope_only() {
    let src = r#"
var globalVar = 100;
function fn() {
    var localVar = 200;
    const indirectEval = eval;
    try {
        return indirectEval("typeof localVar");
    } catch (e) {
        return "undefined";
    }
}
console.log(fn());
"#;
    assert_eq!(run_js(src), vec!["undefined"]);
}

#[test]
fn test_js_direct_eval_creates_local_variable_in_non_strict() {
    let src = r#"
function fn() {
    eval("var newlyCreated = 999;");
    return newlyCreated;
}
console.log(fn());
"#;
    assert_eq!(run_js(src), vec!["999"]);
}

#[test]
fn test_js_direct_eval_strict_mode_isolation_variable_scope() {
    let src = r#"
function fn() {
    "use strict";
    eval("var strictVar = 777;");
    return typeof strictVar;
}
console.log(fn());
"#;
    assert_eq!(run_js(src), vec!["undefined"]);
}

#[test]
fn test_js_indirect_eval_sequence_expression_trigger() {
    let src = r#"
var x = 1;
function fn() {
    var x = 2;
    return (0, eval)("x"); // (0, eval) is an indirect call -> evaluates in global scope!
}
console.log(fn());
"#;
    assert_eq!(run_js(src), vec!["1"]);
}

#[test]
fn test_js_indirect_eval_method_call_trigger() {
    let src = r#"
var g = "Global";
const obj = { eval: eval };
function test() {
    var g = "Local";
    return obj.eval("g");
}
console.log(test());
"#;
    assert_eq!(run_js(src), vec!["Global"]);
}

#[test]
fn test_js_direct_eval_modifies_enclosing_arguments_in_non_strict() {
    let src = r#"
function fn(a) {
    eval("a = 10;");
    return a;
}
console.log(fn(5));
"#;
    assert_eq!(run_js(src), vec!["10"]);
}

#[test]
fn test_js_direct_eval_lexical_environment_chain_traversal() {
    let src = r#"
function outer() {
    const a = 10;
    function inner() {
        const b = 20;
        return eval("a + b");
    }
    return inner();
}
console.log(outer());
"#;
    assert_eq!(run_js(src), vec!["30"]);
}

#[test]
fn test_js_indirect_eval_global_var_declaration() {
    let src = r#"
(0, eval)("var globalFromIndirect = 'CreatedGlobally';");
console.log(globalThis.globalFromIndirect);
"#;
    assert_eq!(run_js(src), vec!["CreatedGlobally"]);
}

#[test]
fn test_js_eval_non_string_argument_returns_argument_as_is() {
    let src = r#"
console.log(eval(123) + "|" + eval(true) + "|" + (eval(null) === null));
"#;
    assert_eq!(run_js(src), vec!["123|true|true"]);
}

#[test]
fn test_js_eval_string_wrapper_object_returns_as_is() {
    let src = r#"
const strObj = new String("2 + 2");
const res = eval(strObj);
console.log(typeof res + "|" + (res === strObj));
"#;
    assert_eq!(run_js(src), vec!["object|true"]); // Primitive string is executed, String object is returned as-is!
}

#[test]
fn test_js_direct_eval_in_arrow_function_lexical_this() {
    let src = r#"
const obj = {
    val: 50,
    getVal: function() {
        const arrow = () => eval("this.val");
        return arrow();
    }
};
console.log(obj.getVal());
"#;
    assert_eq!(run_js(src), vec!["50"]);
}

#[test]
fn test_js_direct_eval_in_constructor_super_call() {
    let src = r#"
class Base { constructor(x) { this.x = x; } }
class Derived extends Base {
    constructor(x) {
        eval("super(x * 2)");
    }
}
console.log(new Derived(10).x);
"#;
    assert_eq!(run_js(src), vec!["20"]);
}

#[test]
fn test_js_eval_completion_value_last_expression() {
    let src = r#"
console.log(eval("1; 2; 3;"));
"#;
    assert_eq!(run_js(src), vec!["3"]);
}

#[test]
fn test_js_eval_empty_string_returns_undefined() {
    let src = r#"
console.log(eval("") === undefined);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_direct_eval_let_and_const_block_scoping() {
    let src = r#"
function fn() {
    eval("let blockScoped = 'Inside';");
    return typeof blockScoped;
}
console.log(fn());
"#;
    assert_eq!(run_js(src), vec!["undefined"]);
}

#[test]
fn test_js_eval_syntax_error_throws_immediately() {
    let src = r#"
try {
    eval("if (true) {");
} catch (e) {
    console.log("Eval SyntaxError: " + (e instanceof SyntaxError));
}
"#;
    assert_eq!(run_js(src), vec!["Eval SyntaxError: true"]);
}

#[test]
fn test_js_direct_eval_access_private_field() {
    let src = r#"
class Secret {
    #code = "1234";
    getCode() {
        return eval("this.#code");
    }
}
console.log(new Secret().getCode());
"#;
    assert_eq!(run_js(src), vec!["1234"]);
}

#[test]
fn test_js_indirect_eval_cannot_access_private_field_throws() {
    let src = r#"
class Secret {
    #code = "1234";
    getCode() {
        return (0, eval)("this.#code");
    }
}
try {
    new Secret().getCode();
} catch (e) {
    console.log("Indirect Eval Private Field SyntaxError");
}
"#;
    assert_eq!(run_js(src), vec!["Indirect Eval Private Field SyntaxError"]);
}

#[test]
fn test_js_eval_alias_called_with_call_or_apply_is_indirect() {
    let src = r#"
var g = "GlobalVal";
function test() {
    var g = "LocalVal";
    return eval.call(null, "g");
}
console.log(test());
"#;
    assert_eq!(run_js(src), vec!["GlobalVal"]);
}

#[test]
fn test_js_eval_reflect_apply_is_indirect() {
    let src = r#"
var g = "GlobalReflect";
function test() {
    var g = "LocalReflect";
    return Reflect.apply(eval, null, ["g"]);
}
console.log(test());
"#;
    assert_eq!(run_js(src), vec!["GlobalReflect"]);
}
