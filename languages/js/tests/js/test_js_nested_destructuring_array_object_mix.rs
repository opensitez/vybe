use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Complex Nested Destructuring (Array & Object Combinations)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_nested_destructuring_object_inside_array() {
    let src = r#"
const [{ id, name }] = [{ id: 1, name: "Alice" }];
console.log(`${id}:${name}`);
"#;
    assert_eq!(run_js(src), vec!["1:Alice"]);
}

#[test]
fn test_js_nested_destructuring_array_inside_object() {
    let src = r#"
const { tags: [firstTag, secondTag] } = { tags: ["js", "rust"] };
console.log(`${firstTag}-${secondTag}`);
"#;
    assert_eq!(run_js(src), vec!["js-rust"]);
}

#[test]
fn test_js_nested_destructuring_deep_object_tree() {
    let src = r#"
const user = {
    profile: {
        personal: { name: "Bob", age: 30 }
    }
};
const { profile: { personal: { name, age } } } = user;
console.log(`${name}|${age}`);
"#;
    assert_eq!(run_js(src), vec!["Bob|30"]);
}

#[test]
fn test_js_nested_destructuring_deep_array_matrix() {
    let src = r#"
const matrix = [[1, 2], [3, 4]];
const [[a, b], [c, d]] = matrix;
console.log(`${a},${b},${c},${d}`);
"#;
    assert_eq!(run_js(src), vec!["1,2,3,4"]);
}

#[test]
fn test_js_nested_destructuring_with_defaults() {
    let src = r#"
const { user: { name = "Anonymous", settings: { theme = "dark" } = {} } = {} } = {};
console.log(`${name}:${theme}`);
"#;
    assert_eq!(run_js(src), vec!["Anonymous:dark"]);
}

#[test]
fn test_js_nested_destructuring_property_alias_inside_nested_object() {
    let src = r#"
const config = { server: { port: 8080 } };
const { server: { port: serverPort } } = config;
console.log(serverPort);
"#;
    assert_eq!(run_js(src), vec!["8080"]);
}

#[test]
fn test_js_nested_destructuring_rest_element_inside_nested_array() {
    let src = r#"
const data = [10, [20, 30, 40]];
const [head, [firstInner, ...restInner]] = data;
console.log(`${head}|${firstInner}|${restInner.join(",")}`);
"#;
    assert_eq!(run_js(src), vec!["10|20|30,40"]);
}

#[test]
fn test_js_nested_destructuring_rest_element_inside_nested_object() {
    let src = r#"
const req = { body: { id: 1, title: "Post", author: "Charlie" } };
const { body: { id, ...attributes } } = req;
console.log(`${id}|${Object.keys(attributes).join(",")}`);
"#;
    assert_eq!(run_js(src), vec!["1|title,author"]);
}

#[test]
fn test_js_nested_destructuring_array_of_objects_elision() {
    let src = r#"
const users = [{ name: "A" }, { name: "B" }, { name: "C" }];
const [, { name: secondUser }] = users;
console.log(secondUser);
"#;
    assert_eq!(run_js(src), vec!["B"]);
}

#[test]
fn test_js_nested_destructuring_object_with_array_default_fallback() {
    let src = r#"
const { items: [first = 0, second = 0] = [10, 20] } = {};
console.log(`${first}:${second}`);
"#;
    assert_eq!(run_js(src), vec!["10:20"]);
}

#[test]
fn test_js_nested_destructuring_computed_property_names() {
    let src = r#"
const key = "data";
const obj = { data: { value: 99 } };
const { [key]: { value } } = obj;
console.log(value);
"#;
    assert_eq!(run_js(src), vec!["99"]);
}

#[test]
fn test_js_nested_destructuring_assignment_existing_variables() {
    let src = r#"
let a, b, c;
[{ a, inner: [b, c] }] = [{ a: 1, inner: [2, 3] }];
console.log(`${a},${b},${c}`);
"#;
    assert_eq!(run_js(src), vec!["1,2,3"]);
}

#[test]
fn test_js_nested_destructuring_null_nested_property_throws() {
    let src = r#"
const obj = { nested: null };
try {
    const { nested: { val } } = obj;
} catch (e) {
    console.log("Nested Null TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Nested Null TypeError"]);
}

#[test]
fn test_js_nested_destructuring_undefined_nested_property_throws() {
    let src = r#"
const obj = {};
try {
    const { nested: { val } } = obj;
} catch (e) {
    console.log("Nested Undefined TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Nested Undefined TypeError"]);
}

#[test]
fn test_js_nested_destructuring_function_return_tuple() {
    let src = r#"
function getResponse() {
    return [200, { data: "SuccessPayload" }];
}
const [status, { data }] = getResponse();
console.log(`${status}:${data}`);
"#;
    assert_eq!(run_js(src), vec!["200:SuccessPayload"]);
}

#[test]
fn test_js_nested_destructuring_symbol_keys_in_nested_objects() {
    let src = r#"
const sym = Symbol("meta");
const obj = { metaContainer: { [sym]: "MetaValue" } };
const { metaContainer: { [sym]: val } } = obj;
console.log(val);
"#;
    assert_eq!(run_js(src), vec!["MetaValue"]);
}

#[test]
fn test_js_nested_destructuring_mixed_primitive_wrappers() {
    let src = r#"
const { str: { length: strLen }, num: { toFixed } } = { str: "abc", num: 5.5 };
console.log(strLen + "|" + (typeof toFixed));
"#;
    assert_eq!(run_js(src), vec!["3|function"]);
}

#[test]
fn test_js_nested_destructuring_default_initializer_side_effects() {
    let src = r#"
let evaluated = false;
const { a: { b = (evaluated = true) } = {} } = { a: { b: "Existing" } };
console.log(b + "|Evaluated=" + evaluated);
"#;
    assert_eq!(run_js(src), vec!["Existing|Evaluated=false"]);
}

#[test]
fn test_js_nested_destructuring_iterable_set_inside_object() {
    let src = r#"
const obj = { numbers: new Set([100, 200]) };
const { numbers: [n1, n2] } = obj;
console.log(`${n1}:${n2}`);
"#;
    assert_eq!(run_js(src), vec!["100:200"]);
}

#[test]
fn test_js_nested_destructuring_four_level_mixed_tree() {
    let src = r#"
const root = [{ a: { b: [{ c: "LeafVal" }] } }];
const [{ a: { b: [{ c: result }] } }] = root;
console.log(result);
"#;
    assert_eq!(run_js(src), vec!["LeafVal"]);
}
