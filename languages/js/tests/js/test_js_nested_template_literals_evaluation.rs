use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Nested Template Literals & Scope Resolution Evaluation
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_nested_template_literal_simple_interpolation() {
    let src = r#"
const outer = "A";
const inner = "B";
console.log(`Outer: ${outer} (Inner: ${`${inner}`})`);
"#;
    assert_eq!(run_js(src), vec!["Outer: A (Inner: B)"]);
}

#[test]
fn test_js_nested_template_literal_conditional_rendering() {
    let src = r#"
const user = { name: "Alice", role: "admin" };
console.log(`User Info: ${user.name} ${user.role ? `[Role: ${user.role.toUpperCase()}]` : ""}`);
"#;
    assert_eq!(run_js(src), vec!["User Info: Alice [Role: ADMIN]"]);
}

#[test]
fn test_js_nested_template_literal_array_map_join() {
    let src = r#"
const items = [{ id: 1 }, { id: 2 }];
const list = `List: ${items.map(item => `Item#${item.id}`).join(", ")}`;
console.log(list);
"#;
    assert_eq!(run_js(src), vec!["List: Item#1, Item#2"]);
}

#[test]
fn test_js_nested_template_literal_three_levels_deep() {
    let src = r#"
const v = 10;
const res = `L1_${`L2_${`L3_${v}`}`}`;
console.log(res);
"#;
    assert_eq!(run_js(src), vec!["L1_L2_L3_10"]);
}

#[test]
fn test_js_nested_template_literal_inside_tagged_template() {
    let src = r#"
function tag(strings, ...values) {
    return strings[0] + values[0];
}
const val = 42;
console.log(tag`Result: ${`Value = ${val}`}`);
"#;
    assert_eq!(run_js(src), vec!["Result: Value = 42"]);
}

#[test]
fn test_js_nested_template_literal_closure_lexical_scope() {
    let src = r#"
function makeFormatter(prefix) {
    return name => `${prefix}: ${`Hello ${name}`}`;
}
const fmt = makeFormatter("LOG");
console.log(fmt("World"));
"#;
    assert_eq!(run_js(src), vec!["LOG: Hello World"]);
}

#[test]
fn test_js_nested_template_literal_object_literal_property() {
    let src = r#"
const obj = {
    msg: `Outer_${`Inner_${100}`}`
};
console.log(obj.msg);
"#;
    assert_eq!(run_js(src), vec!["Outer_Inner_100"]);
}

#[test]
fn test_js_nested_template_literal_expression_evaluation_order() {
    let src = r#"
let order = [];
function first() { order.push(1); return "1"; }
function second() { order.push(2); return "2"; }

const str = `First: ${first()} (${`Second: ${second()}`})`;
console.log(str + "|Order=" + order.join(","));
"#;
    assert_eq!(run_js(src), vec!["First: 1 (Second: 2)|Order=1,2"]);
}

#[test]
fn test_js_nested_template_literal_default_parameter() {
    let src = r#"
function greet(msg = `Default: ${`SubDefault`}`) {
    return msg;
}
console.log(greet());
"#;
    assert_eq!(run_js(src), vec!["Default: SubDefault"]);
}

#[test]
fn test_js_nested_template_literal_switch_case() {
    let src = r#"
const mode = "A";
let out;
switch (mode) {
    case "A":
        out = `Mode: ${`Selected_${mode}`}`;
        break;
}
console.log(out);
"#;
    assert_eq!(run_js(src), vec!["Mode: Selected_A"]);
}

#[test]
fn test_js_nested_template_literal_ternary_chaining() {
    let src = r#"
const score = 85;
const grade = `Grade: ${score >= 90 ? "A" : `${score >= 80 ? `B (${score})` : "C"}`}`;
console.log(grade);
"#;
    assert_eq!(run_js(src), vec!["Grade: B (85)"]);
}

#[test]
fn test_js_nested_template_literal_tagged_inner_template() {
    let src = r#"
function innerTag(strings, val) {
    return strings[0] + val.toUpperCase();
}
console.log(`Outer ${innerTag`Inner ${"text"}`}`);
"#;
    assert_eq!(run_js(src), vec!["Outer Inner TEXT"]);
}

#[test]
fn test_js_nested_template_literal_arrow_body() {
    let src = r#"
const fn = x => `Result: ${`X = ${x}`}`;
console.log(fn(5));
"#;
    assert_eq!(run_js(src), vec!["Result: X = 5"]);
}

#[test]
fn test_js_nested_template_literal_symbol_key_computed_property() {
    let src = r#"
const key = `key_${`1`}`;
const obj = { [key]: "Val1" };
console.log(obj.key_1);
"#;
    assert_eq!(run_js(src), vec!["Val1"]);
}

#[test]
fn test_js_nested_template_literal_try_catch_block() {
    let src = r#"
let res;
try {
    res = `Try_${`Success_${10}`}`;
} catch (e) {
    res = `Catch_${`Error`}`;
}
console.log(res);
"#;
    assert_eq!(run_js(src), vec!["Try_Success_10"]);
}

#[test]
fn test_js_nested_template_literal_multiline_nesting() {
    let src = r#"
const code = `function() {
    return \`${`NestedCode`}\`;
}`;
console.log(code.includes("NestedCode"));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_nested_template_literal_destructuring_assignment() {
    let src = r#"
const { a = `Default_${`1`}` } = {};
console.log(a);
"#;
    assert_eq!(run_js(src), vec!["Default_1"]);
}

#[test]
fn test_js_nested_template_literal_async_await_interpolation() {
    let src = r#"
(async () => {
    const val = await Promise.resolve(50);
    const msg = `Async: ${`Val = ${val}`}`;
    console.log(msg);
})();
"#;
    assert_eq!(run_js(src), vec!["Async: Val = 50"]);
}

#[test]
fn test_js_nested_template_literal_generator_yield_interpolation() {
    let src = r#"
function* gen() {
    yield `Gen_${`Yield_${1}`}`;
}
console.log(gen().next().value);
"#;
    assert_eq!(run_js(src), vec!["Gen_Yield_1"]);
}

#[test]
fn test_js_nested_template_literal_class_private_field_initialization() {
    let src = r#"
class Secret {
    #code = `Priv_${`Secret_${99}`}`;
    getCode() { return this.#code; }
}
console.log(new Secret().getCode());
"#;
    assert_eq!(run_js(src), vec!["Priv_Secret_99"]);
}

#[test]
fn test_js_nested_template_literal_in_iife_expression() {
    let src = r#"
console.log((() => `IIFE_${`Value_${123}`}`)());
"#;
    assert_eq!(run_js(src), vec!["IIFE_Value_123"]);
}

