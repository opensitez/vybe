/// Array higher-order patterns — flatMap, groupBy, partition, zip, chunk
use super::helpers::run_js;

#[test]
fn array_flatmap_flatten_one_level() {
    assert_eq!(
        run_js(
            r#"
const sentences = ["hello world", "foo bar"];
const words = sentences.flatMap(s => s.split(" "));
console.log(words.join(","));
"#
        ),
        vec!["hello,world,foo,bar"]
    );
}

#[test]
fn array_flatmap_filter_and_map() {
    assert_eq!(
        run_js(
            r#"
// flatMap can filter by returning [] for excluded items
const nums = [1, 2, 3, 4, 5];
const evenDoubled = nums.flatMap(x => x % 2 === 0 ? [x * 2] : []);
console.log(evenDoubled.join(","));
"#
        ),
        vec!["4,8"]
    );
}

#[test]
fn partition_pattern() {
    assert_eq!(
        run_js(
            r#"
function partition(arr, pred) {
    return arr.reduce(([pass, fail], x) => {
        return pred(x) ? [[...pass, x], fail] : [pass, [...fail, x]];
    }, [[], []]);
}
const [evens, odds] = partition([1, 2, 3, 4, 5, 6], x => x % 2 === 0);
console.log(evens.join(","));
console.log(odds.join(","));
"#
        ),
        vec!["2,4,6", "1,3,5"]
    );
}

#[test]
fn chunk_array_into_groups() {
    assert_eq!(
        run_js(
            r#"
function chunk(arr, size) {
    const result = [];
    for (let i = 0; i < arr.length; i += size) {
        result.push(arr.slice(i, i + size));
    }
    return result;
}
const chunks = chunk([1, 2, 3, 4, 5, 6, 7], 3);
console.log(chunks.length);
console.log(chunks[0].join(","));
console.log(chunks[1].join(","));
console.log(chunks[2].join(","));
"#
        ),
        vec!["3", "1,2,3", "4,5,6", "7"]
    );
}

#[test]
fn zip_arrays() {
    assert_eq!(
        run_js(
            r#"
function zip(a, b) {
    const len = Math.min(a.length, b.length);
    const out = [];
    for (let i = 0; i < len; i++) {
        out.push([a[i], b[i]]);
    }
    return out;
}
const zipped = zip([1, 2, 3], ["a", "b", "c"], [true, false, true]);
console.log(zipped[0].join(","));
console.log(zipped[1].join(","));
"#
        ),
        vec!["1,a", "2,b"]
    );
}

#[test]
fn flatten_deep_nested() {
    assert_eq!(
        run_js(
            r#"
const nested = [1, [2, [3, [4, [5]]]]];
console.log(nested.flat(Infinity).join(","));
console.log(nested.flat(2).join(","));
"#
        ),
        vec!["1,2,3,4,5", "1,2,3,4,5"]
    );
}

#[test]
fn array_group_by_object() {
    assert_eq!(
        run_js(
            r#"
const items = [
    { type: "fruit", name: "apple" },
    { type: "veggie", name: "carrot" },
    { type: "fruit", name: "banana" },
];
const grouped = Object.groupBy(items, x => x.type);
console.log(grouped.fruit.length);
console.log(grouped.veggie.length);
console.log(grouped.fruit[0].name);
"#
        ),
        vec!["2", "1", "apple"]
    );
}

#[test]
fn array_unique_via_set() {
    assert_eq!(
        run_js(
            r#"
const arr = [1, 2, 2, 3, 1, 4, 3];
const unique = [...new Set(arr)];
console.log(unique.join(","));
"#
        ),
        vec!["1,2,3,4"]
    );
}

#[test]
fn array_rotate() {
    assert_eq!(
        run_js(
            r#"
function rotate(arr, n) {
    const k = ((n % arr.length) + arr.length) % arr.length;
    return [...arr.slice(k), ...arr.slice(0, k)];
}
console.log(rotate([1, 2, 3, 4, 5], 2).join(","));
console.log(rotate([1, 2, 3, 4, 5], 1).join(","));
"#
        ),
        vec!["3,4,5,1,2", "2,3,4,5,1"]
    );
}

#[test]
fn array_sliding_window() {
    assert_eq!(
        run_js(
            r#"
function windows(arr, size) {
    return arr.slice(0, arr.length - size + 1).map((_, i) => arr.slice(i, i + size));
}
const result = windows([1, 2, 3, 4, 5], 3);
console.log(result.length);
console.log(result[0].join(","));
console.log(result[2].join(","));
"#
        ),
        vec!["3", "1,2,3", "3,4,5"]
    );
}

#[test]
fn array_transpose_matrix() {
    assert_eq!(
        run_js(
            r#"
const matrix = [[1, 2, 3], [4, 5, 6], [7, 8, 9]];
const transposed = matrix[0].map((_, col) => matrix.map(row => row[col]));
console.log(transposed[0].join(","));
console.log(transposed[1].join(","));
"#
        ),
        vec!["1,4,7", "2,5,8"]
    );
}

#[test]
fn array_from_with_mapping_function_and_this_arg() {
    assert_eq!(
        run_js(
            r#"
const arr = Array.from([1, 2, 3], v => v * 3);
console.log(arr.join(","));
"#
        ),
        vec!["3,6,9"]
    );
}

#[test]
fn array_methods_skip_missing_holes() {
    assert_eq!(
        run_js(
            r#"
const sparse = [, 2,  4];
const doubled = sparse.map(x => x * 2);
console.log(doubled.length);
console.log(0 in doubled, 1 in doubled, 2 in doubled);
console.log(doubled.join("|"));
"#
        ),
        vec!["4", "false true false", "|4||8"]
    );
}

#[test]
fn array_map_groupby_primitive_keys() {
    assert_eq!(
        run_js(
            r#"
const nums = [1, 2, 3, 4, 5];
const grouped = Map.groupBy(nums, x => x % 2 === 0 ? "even" : "odd");
console.log(grouped.get("even").join(",") + "|" + grouped.get("odd").join(","));
"#
        ),
        vec!["2,4|1,3,5"]
    );
}
