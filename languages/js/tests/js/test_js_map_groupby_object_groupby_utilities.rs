use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: `Map.groupBy` & `Object.groupBy` Helpers (ES2024)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_object_groupby_numeric_or_string_keys() {
    let src = r#"
const numbers = [1, 2, 3, 4, 5, 6];
const result = Object.groupBy(numbers, num => num % 2 === 0 ? "even" : "odd");
console.log(result.even.join(",") + "|" + result.odd.join(","));
"#;
    assert_eq!(run_js(src), vec!["2,4,6|1,3,5"]);
}

#[test]
fn test_js_map_groupby_complex_object_keys() {
    let src = r#"
const keyEven = { type: "even" };
const keyOdd = { type: "odd" };
const numbers = [10, 15, 20, 25];

const map = Map.groupBy(numbers, n => n % 2 === 0 ? keyEven : keyOdd);
console.log(map.get(keyEven).join(",") + "|" + map.get(keyOdd).join(","));
"#;
    assert_eq!(run_js(src), vec!["10,20|15,25"]);
}

#[test]
fn test_js_object_groupby_returns_null_prototype_object() {
    let src = r#"
const items = ["a", "bb", "ccc"];
const grouped = Object.groupBy(items, item => item.length);
console.log(Object.getPrototypeOf(grouped) === null);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_map_groupby_returns_map_instance() {
    let src = r#"
const items = [1, 2];
const grouped = Map.groupBy(items, x => x);
console.log(grouped instanceof Map);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_object_groupby_symbol_keys_coerced_to_string() {
    let src = r#"
const items = ["apple", "apricot", "banana"];
const grouped = Object.groupBy(items, item => item[0]);
console.log(grouped["a"].join(",") + "|" + grouped["b"].join(","));
"#;
    assert_eq!(run_js(src), vec!["apple,apricot|banana"]);
}

#[test]
fn test_js_map_groupby_primitive_keys() {
    let src = r#"
const items = [true, false, true, true];
const grouped = Map.groupBy(items, val => val);
console.log(grouped.get(true).length + "|" + grouped.get(false).length);
"#;
    assert_eq!(run_js(src), vec!["3|1"]);
}

#[test]
fn test_js_object_groupby_callback_index_argument() {
    let src = r#"
const letters = ["a", "b", "c", "d"];
const grouped = Object.groupBy(letters, (char, index) => index < 2 ? "firstHalf" : "secondHalf");
console.log(grouped.firstHalf.join(",") + "|" + grouped.secondHalf.join(","));
"#;
    assert_eq!(run_js(src), vec!["a,b|c,d"]);
}

#[test]
fn test_js_map_groupby_callback_index_argument() {
    let src = r#"
const items = [10, 20, 30, 40];
const grouped = Map.groupBy(items, (item, index) => index % 2);
console.log(grouped.get(0).join(",") + "|" + grouped.get(1).join(","));
"#;
    assert_eq!(run_js(src), vec!["10,30|20,40"]);
}

#[test]
fn test_js_object_groupby_empty_input_array() {
    let src = r#"
const grouped = Object.groupBy([], item => item);
console.log(Object.keys(grouped).length);
"#;
    assert_eq!(run_js(src), vec!["0"]);
}

#[test]
fn test_js_map_groupby_empty_input_iterable() {
    let src = r#"
const grouped = Map.groupBy([], item => item);
console.log(grouped.size);
"#;
    assert_eq!(run_js(src), vec!["0"]);
}

#[test]
fn test_js_object_groupby_symbol_key_throws_typeerror() {
    let src = r#"
try {
    Object.groupBy(["a"], () => Symbol("key"));
} catch (e) {
    console.log("Object.groupBy Symbol Key TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Object.groupBy Symbol Key TypeError"]);
}

#[test]
fn test_js_map_groupby_supports_symbol_keys() {
    let src = r#"
const symKey = Symbol("sym");
const grouped = Map.groupBy(["a", "b"], () => symKey);
console.log(grouped.get(symKey).join(","));
"#;
    assert_eq!(run_js(src), vec!["a,b"]);
}

#[test]
fn test_js_object_groupby_set_iterable() {
    let src = r#"
const set = new Set([1, 2, 3, 4]);
const grouped = Object.groupBy(set, x => x > 2 ? "high" : "low");
console.log(grouped.low.join(",") + "|" + grouped.high.join(","));
"#;
    assert_eq!(run_js(src), vec!["1,2|3,4"]);
}

#[test]
fn test_js_map_groupby_generator_iterable() {
    let src = r#"
function* gen() { yield "cat"; yield "dog"; yield "elephant"; }
const grouped = Map.groupBy(gen(), word => word.length);
console.log(grouped.get(3).join(",") + "|" + grouped.get(8).join(","));
"#;
    assert_eq!(run_js(src), vec!["cat,dog|elephant"]);
}

#[test]
fn test_js_object_groupby_non_callable_callback_throws() {
    let src = r#"
try {
    Object.groupBy([1, 2], "not_a_function");
} catch (e) {
    console.log("Object.groupBy Non-Callable TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Object.groupBy Non-Callable TypeError"]);
}

#[test]
fn test_js_map_groupby_non_callable_callback_throws() {
    let src = r#"
try {
    Map.groupBy([1, 2], null);
} catch (e) {
    console.log("Map.groupBy Non-Callable TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Map.groupBy Non-Callable TypeError"]);
}

#[test]
fn test_js_object_groupby_null_iterable_throws() {
    let src = r#"
try {
    Object.groupBy(null, x => x);
} catch (e) {
    console.log("Object.groupBy Null Iterable TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Object.groupBy Null Iterable TypeError"]);
}

#[test]
fn test_js_map_groupby_preserve_insertion_order_of_keys() {
    let src = r#"
const items = [{ cat: "A", id: 1 }, { cat: "B", id: 2 }, { cat: "A", id: 3 }];
const grouped = Map.groupBy(items, i => i.cat);
console.log([...grouped.keys()].join(","));
"#;
    assert_eq!(run_js(src), vec!["A,B"]);
}

#[test]
fn test_js_object_groupby_sparse_array_holes_visited() {
    let src = r#"
const sparse = [1, , 3];
const grouped = Object.groupBy(sparse, item => item === undefined ? "undef" : "def");
console.log(grouped.def.join(",") + "|countUndef=" + grouped.undef.length);
"#;
    assert_eq!(run_js(src), vec!["1,3|countUndef=1"]);
}

#[test]
fn test_js_map_groupby_nan_key_grouping() {
    let src = r#"
const values = [NaN, 10, NaN, 20];
const grouped = Map.groupBy(values, val => Number.isNaN(val) ? NaN : "number");
console.log(grouped.get(NaN).length + "|" + grouped.get("number").join(","));
"#;
    assert_eq!(run_js(src), vec!["2|10,20"]);
}
