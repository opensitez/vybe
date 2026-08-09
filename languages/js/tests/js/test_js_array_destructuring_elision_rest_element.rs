use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Array Destructuring, Elision (Holes) & Rest Elements
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_array_destructuring_basic_elements() {
    let src = r#"
const [a, b, c] = [10, 20, 30];
console.log(`${a},${b},${c}`);
"#;
    assert_eq!(run_js(src), vec!["10,20,30"]);
}

#[test]
fn test_js_array_destructuring_elision_skipping_elements() {
    let src = r#"
const [first,  third] = [1, 2, 3];
console.log(`${first}|${third}`);
"#;
    assert_eq!(run_js(src), vec!["1|3"]);
}

#[test]
fn test_js_array_destructuring_rest_element() {
    let src = r#"
const [head, ...tail] = [1, 2, 3, 4];
console.log(head + "|" + tail.join(","));
"#;
    assert_eq!(run_js(src), vec!["1|2,3,4"]);
}

#[test]
fn test_js_array_destructuring_default_values() {
    let src = r#"
const [x = 1, y = 2, z = 3] = [10, undefined];
console.log(`${x},${y},${z}`);
"#;
    assert_eq!(run_js(src), vec!["10,2,3"]);
}

#[test]
fn test_js_array_destructuring_sparse_array_holes() {
    let src = r#"
const [a = "defaultA", b = "defaultB"] = [, 20];
console.log(`${a},${b}`);
"#;
    assert_eq!(run_js(src), vec!["defaultA,20"]);
}

#[test]
fn test_js_array_destructuring_iterable_protocol_custom() {
    let src = r#"
const customIterable = {
    *[Symbol.iterator]() {
        yield 100;
        yield 200;
    }
};
const [x, y] = customIterable;
console.log(`${x}:${y}`);
"#;
    assert_eq!(run_js(src), vec!["100:200"]);
}

#[test]
fn test_js_array_destructuring_string_iterable() {
    let src = r#"
const [char1, char2] = "JS";
console.log(`${char1}-${char2}`);
"#;
    assert_eq!(run_js(src), vec!["J-S"]);
}

#[test]
fn test_js_array_destructuring_set_iterable() {
    let src = r#"
const set = new Set([5, 10, 15]);
const [first, second] = set;
console.log(`${first}|${second}`);
"#;
    assert_eq!(run_js(src), vec!["5|10"]);
}

#[test]
fn test_js_array_destructuring_swap_variables() {
    let src = r#"
let a = 1, b = 2;
[a, b] = [b, a];
console.log(`a=${a},b=${b}`);
"#;
    assert_eq!(run_js(src), vec!["a=2,b=1"]);
}

#[test]
fn test_js_array_destructuring_non_iterable_throws_typeerror() {
    let src = r#"
try {
    const [a] = 12345;
} catch (e) {
    console.log("Array Destructure Non-Iterable TypeError");
}
"#;
    assert_eq!(
        run_js(src),
        vec!["Array Destructure Non-Iterable TypeError"]
    );
}

#[test]
fn test_js_array_destructuring_null_target_throws_typeerror() {
    let src = r#"
try {
    const [a] = null;
} catch (e) {
    console.log("Array Destructure Null TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Array Destructure Null TypeError"]);
}

#[test]
fn test_js_array_destructuring_empty_rest_element() {
    let src = r#"
const [a, ...rest] = [10];
console.log(a + "|restLength=" + rest.length);
"#;
    assert_eq!(run_js(src), vec!["10|restLength=0"]);
}

#[test]
fn test_js_array_destructuring_rest_element_must_be_last() {
    let src = r#"
try {
    eval("const [...rest, last] = [1, 2];");
} catch (e) {
    console.log("Rest Element Not Last SyntaxError");
}
"#;
    assert_eq!(run_js(src), vec!["Rest Element Not Last SyntaxError"]);
}

#[test]
fn test_js_array_destructuring_generator_function_lazy_evaluation() {
    let src = r#"
let evaluated = 0;
function* gen() {
    evaluated++; yield 1;
    evaluated++; yield 2;
    evaluated++; yield 3;
}
const [a, b] = gen();
console.log(`${a},${b}|evaluated=${evaluated}`);
"#;
    assert_eq!(run_js(src), vec!["1,2|evaluated=2"]);
}

#[test]
fn test_js_array_destructuring_assignment_to_object_properties() {
    let src = r#"
const obj = {};
[obj.x, obj.y] = [10, 20];
console.log(`${obj.x}:${obj.y}`);
"#;
    assert_eq!(run_js(src), vec!["10:20"]);
}

#[test]
fn test_js_array_destructuring_assignment_to_array_elements() {
    let src = r#"
const arr = [0, 0];
[arr[0], arr[1]] = [5, 15];
console.log(arr.join(","));
"#;
    assert_eq!(run_js(src), vec!["5,15"]);
}

#[test]
fn test_js_array_destructuring_elision_only_pattern() {
    let src = r#"
const [,  c] = [10, 20, 30];
console.log(c);
"#;
    assert_eq!(run_js(src), vec!["30"]);
}

#[test]
fn test_js_array_destructuring_side_effects_in_default_expressions() {
    let src = r#"
let count = 0;
const [a = ++count, b = ++count] = [100];
console.log(`${a},${b}|count=${count}`);
"#;
    assert_eq!(run_js(src), vec!["100,1|count=1"]);
}

#[test]
fn test_js_array_destructuring_map_iterator() {
    let src = r#"
const map = new Map([["k1", "v1"], ["k2", "v2"]]);
const [firstEntry] = map;
console.log(firstEntry.join("="));
"#;
    assert_eq!(run_js(src), vec!["k1=v1"]);
}

#[test]
fn test_js_array_destructuring_excess_target_elements() {
    let src = r#"
const [a, b] = [1, 2, 3, 4, 5];
console.log(`${a},${b}`);
"#;
    assert_eq!(run_js(src), vec!["1,2"]);
}

#[test]
fn test_js_array_destructuring_nested_rest() {
    let src = r#"
const [a, [b, ...subRest]] = [1, [2, 3, 4]];
console.log(a + "|" + b + "|" + subRest.join(","));
"#;
    assert_eq!(run_js(src), vec!["1|2|3,4"]);
}
