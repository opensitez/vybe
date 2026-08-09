use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Comma Operator (`,`), Evaluation Order & Expression Values
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_comma_operator_returns_last_operand_value() {
    let src = r#"
const res = (1, 2, 3);
console.log(res);
"#;
    assert_eq!(run_js(src), vec!["3"]);
}

#[test]
fn test_js_comma_operator_evaluates_all_operands_left_to_right() {
    let src = r#"
const log = [];
const res = (log.push(1), log.push(2), log.push(3), "final");
console.log(res + "|" + log.join(","));
"#;
    assert_eq!(run_js(src), vec!["final|1,2,3"]);
}

#[test]
fn test_js_comma_operator_in_for_loop_head() {
    let src = r#"
const log = [];
for (let i = 0, j = 10; i < 3; i++, j--) {
    log.push(`${i}:${j}`);
}
console.log(log.join("|"));
"#;
    assert_eq!(run_js(src), vec!["0:10|1:9|2:8"]);
}

#[test]
fn test_js_comma_operator_in_return_statement() {
    let src = r#"
let sideEffect = 0;
function fn() {
    return (sideEffect = 100, 42);
}
console.log(fn() + "|SideEffect=" + sideEffect);
"#;
    assert_eq!(run_js(src), vec!["42|SideEffect=100"]);
}

#[test]
fn test_js_comma_operator_indirect_eval_call() {
    let src = r#"
const x = "globalScope";
(() => {
    const x = "localScope";
    const indirectEval = (0, eval); // (0, eval) performs indirect eval in global scope!
    console.log(indirectEval("x"));
})();
"#;
    assert_eq!(run_js(src), vec!["globalScope"]);
}

#[test]
fn test_js_comma_operator_precedence_lowest_of_all_operators() {
    let src = r#"
let a, b;
a = 1, b = 2; // Equivalent to: (a = 1), (b = 2)
console.log(`${a}:${b}`);
"#;
    assert_eq!(run_js(src), vec!["1:2"]);
}

#[test]
fn test_js_comma_operator_vs_argument_separator() {
    let src = r#"
function fn(x) { return x; }
console.log(fn((10, 20))); // Parentheses enforce comma operator, passing 20 to single parameter x!
"#;
    assert_eq!(run_js(src), vec!["20"]);
}

#[test]
fn test_js_comma_operator_vs_array_literal_separator() {
    let src = r#"
const arr = [(1, 2), (3, 4)];
console.log(arr.join(","));
"#;
    assert_eq!(run_js(src), vec!["2,4"]);
}

#[test]
fn test_js_comma_operator_vs_object_literal_separator() {
    let src = r#"
const obj = { a: (10, 20), b: (30, 40) };
console.log(`${obj.a}:${obj.b}`);
"#;
    assert_eq!(run_js(src), vec!["20:40"]);
}

#[test]
fn test_js_comma_operator_in_ternary_condition() {
    let src = r#"
let evaluated = false;
const res = (evaluated = true, false) ? "yes" : "no";
console.log(res + "|Evaluated=" + evaluated);
"#;
    assert_eq!(run_js(src), vec!["no|Evaluated=true"]);
}

#[test]
fn test_js_comma_operator_in_while_loop_condition() {
    let src = r#"
let i = 0;
const log = [];
while (log.push(i), i < 2) {
    i++;
}
console.log(log.join(","));
"#;
    assert_eq!(run_js(src), vec!["0,1,2"]);
}

#[test]
fn test_js_comma_operator_chained_multiple_expressions() {
    let src = r#"
let count = 0;
const val = (count++, count++, count++, count * 10);
console.log(val);
"#;
    assert_eq!(run_js(src), vec!["30"]);
}

#[test]
fn test_js_comma_operator_unwrapping_this_binding() {
    let src = r#"
const obj = {
    val: "ObjVal",
    getVal() { return this ? this.val : "NoThis"; }
};
console.log(obj.getVal() + "|" + (0, obj.getVal)()); // (0, obj.getVal)() invokes method with this = undefined/globalThis!
"#;
    assert_eq!(run_js(src), vec!["ObjVal|NoThis"]);
}

#[test]
fn test_js_comma_operator_with_throw_expression_syntax_error() {
    let src = r#"
try {
    eval("const x = (1, throw new Error());"); // throw statement inside expression is a SyntaxError!
} catch (e) {
    console.log("Comma Operator Throw SyntaxError");
}
"#;
    assert_eq!(run_js(src), vec!["Comma Operator Throw SyntaxError"]);
}

#[test]
fn test_js_comma_operator_in_arrow_function_concise_body() {
    let src = r#"
let count = 0;
const fn = () => (count++, count * 5);
console.log(fn());
"#;
    assert_eq!(run_js(src), vec!["5"]);
}

#[test]
fn test_js_comma_operator_returns_undefined_when_last_operand_is_undefined() {
    let src = r#"
const res = (1, 2, undefined);
console.log(res === undefined);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_comma_operator_with_null_and_boolean_operands() {
    let src = r#"
const res = (null, false, true);
console.log(res);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_comma_operator_with_bigint_and_symbol_operands() {
    let src = r#"
const sym = Symbol("id");
const res = (10n, sym);
console.log(res === sym);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_comma_operator_in_variable_declarations_without_parentheses() {
    let src = r#"
let a = 1, b = 2, c = 3; // Declares three variables, NOT comma operator expression!
console.log(`${a}:${b}:${c}`);
"#;
    assert_eq!(run_js(src), vec!["1:2:3"]);
}

#[test]
fn test_js_comma_operator_completion_value_in_eval() {
    let src = r#"
console.log(eval("10, 20, 30;"));
"#;
    assert_eq!(run_js(src), vec!["30"]);
}

#[test]
fn test_js_comma_operator_exception_aborts_subsequent_operands() {
    let src = r#"
let step = 0;
try {
    const res = (step = 1, (() => { throw new Error("abort"); })(), step = 2);
} catch (e) {
    console.log(e.message + "|step=" + step);
}
"#;
    assert_eq!(run_js(src), vec!["abort|step=1"]);
}
