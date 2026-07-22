use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Map & Set Iteration Mechanics (keys, values, entries, forEach)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_map_keys_values_entries_iterators() {
    let src = r#"
const map = new Map([["a", 1], ["b", 2]]);
console.log([...map.keys()].join(","));
console.log([...map.values()].join(","));
console.log([...map.entries()].map(e => e.join("=")).join(","));
"#;
    assert_eq!(run_js(src), vec!["a,b", "1,2", "a=1,b=2"]);
}

#[test]
fn test_js_set_keys_values_entries_iterators() {
    let src = r#"
const set = new Set(["x", "y"]);
console.log([...set.keys()].join(","));
console.log([...set.values()].join(","));
console.log([...set.entries()].map(e => e.join("=")).join(","));
"#;
    assert_eq!(run_js(src), vec!["x,y", "x,y", "x=x,y=y"]);
}

#[test]
fn test_js_map_for_of_iteration_yields_entry_tuples() {
    let src = r#"
const map = new Map([["k1", 10], ["k2", 20]]);
const res = [];
for (const [k, v] of map) {
    res.push(`${k}:${v}`);
}
console.log(res.join(","));
"#;
    assert_eq!(run_js(src), vec!["k1:10,k2:20"]);
}

#[test]
fn test_js_set_for_of_iteration_yields_elements() {
    let src = r#"
const set = new Set([10, 20, 30]);
const res = [];
for (const val of set) {
    res.push(val);
}
console.log(res.join(","));
"#;
    assert_eq!(run_js(src), vec!["10,20,30"]);
}

#[test]
fn test_js_map_foreach_callback_arguments() {
    let src = r#"
const map = new Map([["a", 100]]);
map.forEach((val, key, m) => {
    console.log(`${key}=${val}|isMap=${m === map}`);
});
"#;
    assert_eq!(run_js(src), vec!["a=100|isMap=true"]);
}

#[test]
fn test_js_set_foreach_callback_arguments() {
    let src = r#"
const set = new Set(["elem"]);
set.forEach((val1, val2, s) => {
    console.log(`${val1}:${val2}|isSet=${s === set}`); // In Set forEach, first and second args are identical element!
});
"#;
    assert_eq!(run_js(src), vec!["elem:elem|isSet=true"]);
}

#[test]
fn test_js_map_foreach_this_arg() {
    let src = r#"
const context = { prefix: "Item" };
const map = new Map([["1", "A"]]);
map.forEach(function(val, key) {
    console.log(`${this.prefix}:${key}->${val}`);
}, context);
"#;
    assert_eq!(run_js(src), vec!["Item:1->A"]);
}

#[test]
fn test_js_map_insertion_order_preservation() {
    let src = r#"
const map = new Map();
map.set("c", 3);
map.set("a", 1);
map.set("b", 2);
console.log([...map.keys()].join(","));
"#;
    assert_eq!(run_js(src), vec!["c,a,b"]);
}

#[test]
fn test_js_map_mutation_during_iteration_adds_visited() {
    let src = r#"
const map = new Map([["a", 1]]);
const visited = [];
for (const [k, v] of map) {
    visited.push(k);
    if (k === "a") map.set("b", 2);
}
console.log(visited.join(","));
"#;
    assert_eq!(run_js(src), vec!["a,b"]);
}

#[test]
fn test_js_map_mutation_during_iteration_deletes_unvisited() {
    let src = r#"
const map = new Map([["a", 1], ["b", 2], ["c", 3]]);
const visited = [];
for (const [k, v] of map) {
    visited.push(k);
    if (k === "a") map.delete("b");
}
console.log(visited.join(","));
"#;
    assert_eq!(run_js(src), vec!["a,c"]);
}

#[test]
fn test_js_set_mutation_during_iteration_adds_visited() {
    let src = r#"
const set = new Set([1]);
const visited = [];
for (const val of set) {
    visited.push(val);
    if (val === 1) set.add(2);
}
console.log(visited.join(","));
"#;
    assert_eq!(run_js(src), vec!["1,2"]);
}

#[test]
fn test_js_set_mutation_during_iteration_deletes_unvisited() {
    let src = r#"
const set = new Set([1, 2, 3]);
const visited = [];
for (const val of set) {
    visited.push(val);
    if (val === 1) set.delete(2);
}
console.log(visited.join(","));
"#;
    assert_eq!(run_js(src), vec!["1,3"]);
}

#[test]
fn test_js_map_default_iterator_equals_entries() {
    let src = r#"
const map = new Map([["x", 9]]);
console.log(map[Symbol.iterator] === map.entries);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_set_default_iterator_equals_values() {
    let src = r#"
const set = new Set([9]);
console.log(set[Symbol.iterator] === set.values);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_map_iterator_next_return_structure() {
    let src = r#"
const map = new Map([["key", "val"]]);
const iter = map.entries();
const step1 = iter.next();
const step2 = iter.next();

console.log(`${step1.value.join("=")}|done=${step1.done}`);
console.log(`${step2.value}|done=${step2.done}`);
"#;
    assert_eq!(
        run_js(src),
        vec!["key=val|done=false", "undefined|done=true"]
    );
}

#[test]
fn test_js_set_iterator_next_return_structure() {
    let src = r#"
const set = new Set([42]);
const iter = set.values();
const step1 = iter.next();
const step2 = iter.next();

console.log(`${step1.value}|done=${step1.done}`);
console.log(`${step2.value}|done=${step2.done}`);
"#;
    assert_eq!(run_js(src), vec!["42|done=false", "undefined|done=true"]);
}

#[test]
fn test_js_map_empty_collection_iteration() {
    let src = r#"
const map = new Map();
let count = 0;
for (const _ of map) count++;
console.log(count);
"#;
    assert_eq!(run_js(src), vec!["0"]);
}

#[test]
fn test_js_set_empty_collection_iteration() {
    let src = r#"
const set = new Set();
let count = 0;
for (const _ of set) count++;
console.log(count);
"#;
    assert_eq!(run_js(src), vec!["0"]);
}

#[test]
fn test_js_map_iterator_prototype_chain() {
    let src = r#"
const map = new Map();
const iter = map.keys();
const proto = Object.getPrototypeOf(iter);
console.log(typeof iter[Symbol.iterator] === "function");
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_set_array_from_mapping_function() {
    let src = r#"
const set = new Set([1, 2, 3]);
const doubled = Array.from(set, x => x * 2);
console.log(doubled.join(","));
"#;
    assert_eq!(run_js(src), vec!["2,4,6"]);
}
