use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Array Iteration Methods (map, filter, reduce, reduceRight)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_array_map_basic_transformation() {
    let src = r#"
const nums = [1, 2, 3];
const doubled = nums.map(x => x * 2);
console.log(doubled.join(","));
"#;
    assert_eq!(run_js(src), vec!["2,4,6"]);
}

#[test]
fn test_js_array_map_index_and_array_arguments() {
    let src = r#"
const letters = ["a", "b"];
const res = letters.map((val, idx, arr) => `${val}:${idx}:${arr.length}`);
console.log(res.join("|"));
"#;
    assert_eq!(run_js(src), vec!["a:0:2|b:1:2"]);
}

#[test]
fn test_js_array_map_this_arg_binding() {
    let src = r#"
const context = { multiplier: 10 };
const nums = [1, 2];
const res = nums.map(function(x) { return x * this.multiplier; }, context);
console.log(res.join(","));
"#;
    assert_eq!(run_js(src), vec!["10,20"]);
}

#[test]
fn test_js_array_filter_predicate() {
    let src = r#"
const nums = [1, 2, 3, 4, 5];
const evens = nums.filter(x => x % 2 === 0);
console.log(evens.join(","));
"#;
    assert_eq!(run_js(src), vec!["2,4"]);
}

#[test]
fn test_js_array_reduce_sum_with_initial_value() {
    let src = r#"
const nums = [10, 20, 30];
const sum = nums.reduce((acc, curr) => acc + curr, 100);
console.log(sum);
"#;
    assert_eq!(run_js(src), vec!["160"]);
}

#[test]
fn test_js_array_reduce_sum_without_initial_value() {
    let src = r#"
const nums = [10, 20, 30];
const sum = nums.reduce((acc, curr) => acc + curr);
console.log(sum);
"#;
    assert_eq!(run_js(src), vec!["60"]);
}

#[test]
fn test_js_array_reduce_right_right_to_left_order() {
    let src = r#"
const words = ["a", "b", "c"];
const str = words.reduceRight((acc, curr) => acc + curr, "");
console.log(str);
"#;
    assert_eq!(run_js(src), vec!["cba"]);
}

#[test]
fn test_js_array_reduce_empty_array_no_initial_value_throws_typeerror() {
    let src = r#"
try {
    [].reduce((acc, curr) => acc + curr);
} catch (e) {
    console.log("Reduce Empty Array TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Reduce Empty Array TypeError"]);
}

#[test]
fn test_js_array_reduce_single_element_no_initial_value_returns_element() {
    let src = r#"
const val = [42].reduce((acc, curr) => acc + curr);
console.log(val);
"#;
    assert_eq!(run_js(src), vec!["42"]);
}

#[test]
fn test_js_array_map_sparse_array_holes_preserved() {
    let src = r#"
const sparse = [1, , 3];
const mapped = sparse.map(x => x * 2);
console.log(mapped.length + "|hasHole=" + !(1 in mapped));
"#;
    assert_eq!(run_js(src), vec!["3|hasHole=true"]);
}

#[test]
fn test_js_array_filter_sparse_array_holes_skipped() {
    let src = r#"
const sparse = [1, , 3];
const filtered = sparse.filter(() => true);
console.log(filtered.length + "|" + filtered.join(","));
"#;
    assert_eq!(run_js(src), vec!["2|1,3"]);
}

#[test]
fn test_js_array_reduce_sparse_array_holes_skipped() {
    let src = r#"
const sparse = [1, , 3];
const sum = sparse.reduce((acc, x) => acc + x, 0);
console.log(sum);
"#;
    assert_eq!(run_js(src), vec!["4"]);
}

#[test]
fn test_js_array_map_non_callable_callback_throws() {
    let src = r#"
try {
    [1, 2].map("not_fn");
} catch (e) {
    console.log("Map Non-Callable TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Map Non-Callable TypeError"]);
}

#[test]
fn test_js_array_reduce_flatten_matrix() {
    let src = r#"
const matrix = [[1, 2], [3, 4]];
const flat = matrix.reduce((acc, row) => acc.concat(row), []);
console.log(flat.join(","));
"#;
    assert_eq!(run_js(src), vec!["1,2,3,4"]);
}

#[test]
fn test_js_array_reduce_histogram_builder() {
    let src = r#"
const fruits = ["apple", "banana", "apple", "orange", "banana", "apple"];
const counts = fruits.reduce((acc, fruit) => {
    acc[fruit] = (acc[fruit] || 0) + 1;
    return acc;
}, {});
console.log(`${counts.apple}:${counts.banana}:${counts.orange}`);
"#;
    assert_eq!(run_js(src), vec!["3:2:1"]);
}

#[test]
fn test_js_array_map_mutation_during_iteration() {
    let src = r#"
const arr = [1, 2, 3];
const res = arr.map((x, idx, a) => {
    if (idx === 0) a[2] = 99; // Mutates original array element before visited
    return x;
});
console.log(res.join(","));
"#;
    assert_eq!(run_js(src), vec!["1,2,99"]);
}

#[test]
fn test_js_array_filter_subclass_species() {
    let src = r#"
class CustomArray extends Array {}
const ca = new CustomArray(1, 2, 3, 4);
const filtered = ca.filter(x => x > 2);
console.log(filtered.join(",") + "|isCustom=" + (filtered instanceof CustomArray));
"#;
    assert_eq!(run_js(src), vec!["3,4|isCustom=true"]);
}

#[test]
fn test_js_array_map_chaining() {
    let src = r#"
const nums = [1, 2, 3, 4, 5];
const res = nums.filter(x => x % 2 !== 0).map(x => x * 10);
console.log(res.join(","));
"#;
    assert_eq!(run_js(src), vec!["10,30,50"]);
}

#[test]
fn test_js_array_reduce_right_initial_value_omitted_empty_sparse() {
    let src = r#"
const sparse = [, , 42, ,];
const val = sparse.reduceRight((acc, x) => acc + x);
console.log(val);
"#;
    assert_eq!(run_js(src), vec!["42"]);
}

#[test]
fn test_js_array_map_length_fixed_at_start() {
    let src = r#"
const arr = [1, 2];
const res = arr.map((x, idx, a) => {
    if (idx === 0) a.push(3); // Pushed element should NOT be visited by map!
    return x * 10;
});
console.log(res.join(",") + "|arrLength=" + arr.length);
"#;
    assert_eq!(run_js(src), vec!["10,20|arrLength=3"]);
}
