use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: `Object.groupBy` & `Map.groupBy` Aggregation Utilities (ES2024)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_object_groupby_string_keys() {
    let src = r#"
const inventory = [
    { name: "asparagus", type: "vegetable" },
    { name: "bananas", type: "fruit" },
    { name: "goat", type: "meat" },
    { name: "cherries", type: "fruit" }
];
const result = Object.groupBy(inventory, item => item.type);
console.log(`${result.vegetable.length}:${result.fruit.length}:${result.meat.length}`);
"#;
    assert_eq!(run_js(src), vec!["1:2:1"]);
}

#[test]
fn test_js_map_groupby_object_keys() {
    let src = r#"
const rest1 = { name: "RestA" };
const rest2 = { name: "RestB" };
const goods = [
    { item: "apple", rest: rest1 },
    { item: "pear", rest: rest1 },
    { item: "steak", rest: rest2 }
];
const groupedMap = Map.groupBy(goods, g => g.rest);
console.log((groupedMap instanceof Map) + "|" + groupedMap.get(rest1).length + "|" + groupedMap.get(rest2).length);
"#;
    assert_eq!(run_js(src), vec!["true|2|1"]);
}

#[test]
fn test_js_object_groupby_returns_null_proto_object() {
    let src = r#"
const result = Object.groupBy([1, 2, 3], x => x % 2 === 0 ? "even" : "odd");
console.log(Object.getPrototypeOf(result) === null);
"#;
    assert_eq!(run_js(src), vec!["true"]); // Object.groupBy returns a null-prototype object!
}

#[test]
fn test_js_object_groupby_coerces_keys_to_strings() {
    let src = r#"
const numbers = [10, 20, 30];
const grouped = Object.groupBy(numbers, x => x);
console.log(Object.keys(grouped).join(","));
"#;
    assert_eq!(run_js(src), vec!["10,20,30"]);
}

#[test]
fn test_js_map_groupby_preserves_primitive_key_types() {
    let src = r#"
const items = [1, "1", true, 1n];
const grouped = Map.groupBy(items, x => x);
console.log(grouped.size + "|" + (typeof [...grouped.keys()][0]));
"#;
    assert_eq!(run_js(src), vec!["4|number"]); // Map.groupBy keeps exact key types (number, string, boolean, bigint)!
}

#[test]
fn test_js_object_groupby_callback_arguments() {
    let src = r#"
const arr = ["a", "b"];
const log = [];
Object.groupBy(arr, (val, index) => {
    log.push(`${val}:${index}`);
    return "group";
});
console.log(log.join("|"));
"#;
    assert_eq!(run_js(src), vec!["a:0|b:1"]);
}

#[test]
fn test_js_map_groupby_symbol_keys() {
    let src = r#"
const s1 = Symbol("group1");
const s2 = Symbol("group2");
const arr = [1, 2, 3, 4];
const grouped = Map.groupBy(arr, x => x % 2 === 0 ? s1 : s2);
console.log(grouped.get(s1).join(",") + "|" + grouped.get(s2).join(","));
"#;
    assert_eq!(run_js(src), vec!["2,4|1,3"]);
}

#[test]
fn test_js_object_groupby_empty_iterable() {
    let src = r#"
const grouped = Object.groupBy([], () => "key");
console.log(Object.keys(grouped).length);
"#;
    assert_eq!(run_js(src), vec!["0"]);
}

#[test]
fn test_js_map_groupby_empty_iterable() {
    let src = r#"
const grouped = Map.groupBy([], () => "key");
console.log(grouped.size);
"#;
    assert_eq!(run_js(src), vec!["0"]);
}

#[test]
fn test_js_object_groupby_symbol_key_throws_typeerror() {
    let src = r#"
try {
    Object.groupBy([1], () => Symbol("sym")); // Object.groupBy coerces key to string, Symbol throws TypeError!
} catch (e) {
    console.log("Object.groupBy Symbol Key TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Object.groupBy Symbol Key TypeError"]);
}

#[test]
fn test_js_object_groupby_non_callable_callback_throws_typeerror() {
    let src = r#"
try {
    Object.groupBy([1, 2], "not_a_fn");
} catch (e) {
    console.log("Object.groupBy Callback TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Object.groupBy Callback TypeError"]);
}

#[test]
fn test_js_map_groupby_non_callable_callback_throws_typeerror() {
    let src = r#"
try {
    Map.groupBy([1, 2], null);
} catch (e) {
    console.log("Map.groupBy Callback TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Map.groupBy Callback TypeError"]);
}

#[test]
fn test_js_object_groupby_custom_iterable_source() {
    let src = r#"
const set = new Set(["apple", "banana", "avocado"]);
const grouped = Object.groupBy(set, word => word[0]);
console.log(grouped.a.join(",") + "|" + grouped.b.join(","));
"#;
    assert_eq!(run_js(src), vec!["apple,avocado|banana"]);
}

#[test]
fn test_js_map_groupby_custom_iterable_source() {
    let src = r#"
const generator = function*() { yield 10; yield 20; yield 15; };
const grouped = Map.groupBy(generator(), x => x >= 15);
console.log(grouped.get(true).join(",") + "|" + grouped.get(false).join(","));
"#;
    assert_eq!(run_js(src), vec!["20,15|10"]);
}

#[test]
fn test_js_object_groupby_numeric_indices_ordering() {
    let src = r#"
const arr = [100, 200, 300];
const result = Object.groupBy(arr, (val, idx) => idx);
console.log(result[0][0] + "|" + result[1][0]);
"#;
    assert_eq!(run_js(src), vec!["100|200"]);
}

#[test]
fn test_js_map_groupby_undefined_and_null_keys() {
    let src = r#"
const arr = [1, 2, 3];
const grouped = Map.groupBy(arr, x => x === 1 ? null : undefined);
console.log(grouped.get(null).length + "|" + grouped.get(undefined).length);
"#;
    assert_eq!(run_js(src), vec!["1|2"]);
}

#[test]
fn test_js_object_groupby_undefined_and_null_keys_coerced_to_strings() {
    let src = r#"
const arr = [1, 2];
const grouped = Object.groupBy(arr, x => x === 1 ? null : undefined);
console.log(grouped["null"].length + "|" + grouped["undefined"].length);
"#;
    assert_eq!(run_js(src), vec!["1|1"]);
}

#[test]
fn test_js_object_groupby_sparse_array_holes_visited() {
    let src = r#"
const sparse = [1,  3];
const grouped = Object.groupBy(sparse, x => typeof x);
console.log(grouped.number.length + "|" + grouped.undefined.length);
"#;
    assert_eq!(run_js(src), vec!["2|1"]);
}

#[test]
fn test_js_map_groupby_nan_keys_grouped_together() {
    let src = r#"
const arr = [1, 2, 3];
const grouped = Map.groupBy(arr, () => NaN);
console.log(grouped.size + "|" + grouped.get(NaN).length); // NaN key uses SameValueZero equality!
"#;
    assert_eq!(run_js(src), vec!["1|3"]);
}

#[test]
fn test_js_object_groupby_non_iterable_source_throws_typeerror() {
    let src = r#"
try {
    Object.groupBy(12345, () => "group");
} catch (e) {
    console.log("Object.groupBy Non-Iterable TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Object.groupBy Non-Iterable TypeError"]);
}

#[test]
fn test_js_map_groupby_non_iterable_source_throws_typeerror() {
    let src = r#"
try {
    Map.groupBy(null, () => "group");
} catch (e) {
    console.log("Map.groupBy Non-Iterable TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Map.groupBy Non-Iterable TypeError"]);
}

