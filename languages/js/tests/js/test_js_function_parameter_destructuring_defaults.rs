use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Function Parameter Destructuring & Default Value Binding
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_parameter_destructuring_object() {
    let src = r#"
function printUser({ name, age }) {
    console.log(`${name}:${age}`);
}
printUser({ name: "Alice", age: 30 });
"#;
    assert_eq!(run_js(src), vec!["Alice:30"]);
}

#[test]
fn test_js_parameter_destructuring_array() {
    let src = r#"
function printCoords([x, y]) {
    console.log(`${x},${y}`);
}
printCoords([100, 200]);
"#;
    assert_eq!(run_js(src), vec!["100,200"]);
}

#[test]
fn test_js_parameter_destructuring_object_defaults() {
    let src = r#"
function configure({ port = 8080, host = "localhost" } = {}) {
    console.log(`${host}:${port}`);
}
configure({ port: 3000 });
configure();
"#;
    assert_eq!(run_js(src), vec!["localhost:3000", "localhost:8080"]);
}

#[test]
fn test_js_parameter_destructuring_array_defaults() {
    let src = r#"
function processRange([min = 0, max = 100] = []) {
    console.log(`${min}->${max}`);
}
processRange([10]);
processRange();
"#;
    assert_eq!(run_js(src), vec!["10->100", "0->100"]);
}

#[test]
fn test_js_parameter_destructuring_alias_and_default() {
    let src = r#"
function setLimits({ max: limit = 50 } = {}) {
    console.log(limit);
}
setLimits({ max: 99 });
setLimits({});
"#;
    assert_eq!(run_js(src), vec!["99", "50"]);
}

#[test]
fn test_js_parameter_destructuring_rest_parameter() {
    let src = r#"
function sumAll([head, ...tail]) {
    return head + tail.reduce((a, b) => a + b, 0);
}
console.log(sumAll([10, 1, 2, 3]));
"#;
    assert_eq!(run_js(src), vec!["16"]);
}

#[test]
fn test_js_parameter_destructuring_nested_object_parameters() {
    let src = r#"
function render({ data: { items: [first] } }) {
    console.log(first);
}
render({ data: { items: ["Item1", "Item2"] } });
"#;
    assert_eq!(run_js(src), vec!["Item1"]);
}

#[test]
fn test_js_parameter_destructuring_function_length_property() {
    let src = r#"
function f1(a, { b }, c = 1) {}
function f2({ a }, b = 2, c) {}
console.log(f1.length + "|" + f2.length); // Parameters up to first default/destructured without outer fallback
"#;
    assert_eq!(run_js(src), vec!["1|1"]);
}

#[test]
fn test_js_parameter_destructuring_missing_argument_without_default_throws() {
    let src = r#"
function required({ prop }) {}
try {
    required();
} catch (e) {
    console.log("Parameter Destructure Undefined TypeError");
}
"#;
    assert_eq!(
        run_js(src),
        vec!["Parameter Destructure Undefined TypeError"]
    );
}

#[test]
fn test_js_parameter_destructuring_arrow_function() {
    let src = r#"
const getFull = ({ first, last }) => `${first} ${last}`;
console.log(getFull({ first: "John", last: "Doe" }));
"#;
    assert_eq!(run_js(src), vec!["John Doe"]);
}

#[test]
fn test_js_parameter_destructuring_method_definition() {
    let src = r#"
const service = {
    handle({ id, status = "ok" }) {
        return `${id}:${status}`;
    }
};
console.log(service.handle({ id: 99 }));
"#;
    assert_eq!(run_js(src), vec!["99:ok"]);
}

#[test]
fn test_js_parameter_destructuring_default_expression_scope() {
    let src = r#"
const defaultPort = 80;
function connect({ port = defaultPort } = {}) {
    return port;
}
console.log(connect());
"#;
    assert_eq!(run_js(src), vec!["80"]);
}

#[test]
fn test_js_parameter_destructuring_earlier_parameter_reference_in_default() {
    let src = r#"
function compute(width, { height = width * 2 } = {}) {
    return width * height;
}
console.log(compute(10));
"#;
    assert_eq!(run_js(src), vec!["200"]);
}

#[test]
fn test_js_parameter_destructuring_side_effects_in_default_initializer() {
    let src = r#"
let calls = 0;
function log({ val = ++calls } = {}) {
    return val;
}
console.log(log({ val: 100 }) + "|" + log() + "|Calls=" + calls);
"#;
    assert_eq!(run_js(src), vec!["100|1|Calls=1"]);
}

#[test]
fn test_js_parameter_destructuring_async_function() {
    let src = r#"
async function fetchData({ url, timeout = 1000 }) {
    await Promise.resolve();
    return `${url}:${timeout}`;
}
fetchData({ url: "https://api.com" }).then(res => console.log(res));
"#;
    assert_eq!(run_js(src), vec!["https://api.com:1000"]);
}

#[test]
fn test_js_parameter_destructuring_generator_function() {
    let src = r#"
function* step({ count }) {
    for (let i = 1; i <= count; i++) yield i;
}
console.log([...step({ count: 3 })].join(","));
"#;
    assert_eq!(run_js(src), vec!["1,2,3"]);
}

#[test]
fn test_js_parameter_destructuring_class_constructor() {
    let src = r#"
class Config {
    constructor({ env = "dev", debug = false } = {}) {
        this.env = env;
        this.debug = debug;
    }
}
const cfg = new Config();
console.log(cfg.env + "|" + cfg.debug);
"#;
    assert_eq!(run_js(src), vec!["dev|false"]);
}

#[test]
fn test_js_parameter_destructuring_computed_property_parameter() {
    let src = r#"
const propName = "key";
function extract({ [propName]: val }) {
    return val;
}
console.log(extract({ key: "ExtractedValue" }));
"#;
    assert_eq!(run_js(src), vec!["ExtractedValue"]);
}

#[test]
fn test_js_parameter_destructuring_null_argument_overrides_outer_default() {
    let src = r#"
function test(opts = { val: 10 }) {
    return opts;
}
console.log(test(null) === null);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_parameter_destructuring_arguments_object_binding() {
    let src = r#"
function testArgs({ a, b }) {
    console.log(arguments[0].a + "|" + arguments[0].b);
}
testArgs({ a: 1, b: 2 });
"#;
    assert_eq!(run_js(src), vec!["1|2"]);
}

#[test]
fn test_js_parameter_destructuring_array_from_string() {
    let src = r#"
function f([a, b, c]) {
    return `${a}-${b}-${c}`;
}
console.log(f("XYZ"));
"#;
    assert_eq!(run_js(src), vec!["X-Y-Z"]);
}
