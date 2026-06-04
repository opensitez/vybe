/// ES2022/ES2023 Array methods — tests for features not covered by
/// test_ecma_arrays.rs or test_string_array_advanced.rs.
///
/// Covers: at(), findLast/findLastIndex, toSorted/toReversed/toSpliced/with,
/// flat (depth), flatMap with index, Array.from (Set/Map/mapper), Array.of,
/// groupBy via reduce, fill/copyWithin variants, iterators, indexOf/lastIndexOf
/// with fromIndex, chaining, spread copy, findIndex not-found.
use super::helpers::run_js;

// ===================================================================
// 1. Array.prototype.at() — positive index
// ===================================================================

#[test]
fn array_at_positive_index() {
    assert_eq!(
        run_js(
            r#"
const arr = ["a", "b", "c", "d"];
console.log(arr.at(0));
console.log(arr.at(2));
"#
        ),
        vec!["a", "c"]
    );
}

// ===================================================================
// 2. Array.prototype.at() — negative index (-1 gets last)
// ===================================================================

#[test]
fn array_at_negative_index_last() {
    assert_eq!(
        run_js(
            r#"
const arr = [10, 20, 30, 40, 50];
console.log(arr.at(-1));
console.log(arr.at(-3));
"#
        ),
        vec!["50", "30"]
    );
}

// ===================================================================
// 3. Array.prototype.findLast()
// ===================================================================

#[test]
fn array_findlast_basic() {
    assert_eq!(
        run_js(
            r#"
const arr = [1, 3, 5, 7, 2, 4, 6];
const last = arr.findLast(x => x % 2 === 0);
console.log(last);
"#
        ),
        vec!["6"]
    );
}

// ===================================================================
// 4. Array.prototype.findLastIndex()
// ===================================================================

#[test]
fn array_findlastindex_basic() {
    assert_eq!(
        run_js(
            r#"
const arr = [1, 2, 3, 4, 5, 6];
const idx = arr.findLastIndex(x => x % 2 === 0);
console.log(idx);
"#
        ),
        vec!["5"]
    );
}

// ===================================================================
// 5. Array.prototype.toSorted() — returns sorted copy, original untouched
// ===================================================================

#[test]
fn array_tosorted_does_not_mutate() {
    assert_eq!(
        run_js(
            r#"
const orig = [3, 1, 4, 1, 5];
const sorted = orig.toSorted((a, b) => a - b);
console.log(orig.join(","));
console.log(sorted.join(","));
"#
        ),
        vec!["3,1,4,1,5", "1,1,3,4,5"]
    );
}

// ===================================================================
// 6. Array.prototype.toReversed() — returns reversed copy, original untouched
// ===================================================================

#[test]
fn array_toreversed_does_not_mutate() {
    assert_eq!(
        run_js(
            r#"
const orig = [1, 2, 3, 4, 5];
const rev = orig.toReversed();
console.log(orig.join(","));
console.log(rev.join(","));
"#
        ),
        vec!["1,2,3,4,5", "5,4,3,2,1"]
    );
}

// ===================================================================
// 7. Array.prototype.toSpliced() — returns copy with elements replaced
// ===================================================================

#[test]
fn array_tospliced_returns_new_array() {
    assert_eq!(
        run_js(
            r#"
const orig = [1, 2, 3, 4, 5];
const spliced = orig.toSpliced(1, 2, 10, 20);
console.log(orig.join(","));
console.log(spliced.join(","));
"#
        ),
        vec!["1,2,3,4,5", "1,10,20,4,5"]
    );
}

// ===================================================================
// 8. Array.prototype.with() — returns copy with one element replaced
// ===================================================================

#[test]
fn array_with_replaces_single_element() {
    assert_eq!(
        run_js(
            r#"
const orig = [1, 2, 3, 4, 5];
const updated = orig.with(2, 99);
console.log(orig.join(","));
console.log(updated.join(","));
"#
        ),
        vec!["1,2,3,4,5", "1,2,99,4,5"]
    );
}

// ===================================================================
// 9. Array.prototype.flat() — depth 1
// ===================================================================

#[test]
fn array_flat_depth_one() {
    assert_eq!(
        run_js(
            r#"
const arr = [1, [2, 3], [4, [5, 6]]];
console.log(arr.flat(1).join(","));
"#
        ),
        vec!["1,2,3,4,5,6"]
    );
}

// ===================================================================
// 10. Array.prototype.flat() — depth 2
// ===================================================================

#[test]
fn array_flat_depth_two() {
    assert_eq!(
        run_js(
            r#"
const arr = [1, [2, [3, [4]]]];
console.log(arr.flat(2).join(","));
"#
        ),
        vec!["1,2,3,4"]
    );
}

// ===================================================================
// 11. Array.prototype.flat() — Infinity depth
// ===================================================================

#[test]
fn array_flat_infinity_depth() {
    assert_eq!(
        run_js(
            r#"
const arr = [[1, 2], [3, [4, 5]]];
console.log(arr.flat(Infinity).join(","));
"#
        ),
        vec!["1,2,3,4,5"]
    );
}

// ===================================================================
// 12. Array.prototype.flatMap() — maps then flattens one level
// ===================================================================

#[test]
fn array_flatmap_maps_and_flattens() {
    assert_eq!(
        run_js(
            r#"
const arr = ["hello world", "foo bar baz"];
const words = arr.flatMap(s => s.split(" "));
console.log(words.length);
console.log(words.join(","));
"#
        ),
        vec!["5", "hello,world,foo,bar,baz"]
    );
}

// ===================================================================
// 13. Array.from() — with mapping function
// ===================================================================

#[test]
fn array_from_with_mapping_function() {
    assert_eq!(
        run_js(
            r#"
const arr = Array.from([1, 2, 3, 4], x => x * x);
console.log(arr.join(","));
"#
        ),
        vec!["1,4,9,16"]
    );
}

// ===================================================================
// 14. Array.from() — on string
// ===================================================================

#[test]
fn array_from_on_string() {
    assert_eq!(
        run_js(
            r#"
const arr = Array.from("rust");
console.log(arr.join("-"));
"#
        ),
        vec!["r-u-s-t"]
    );
}

// ===================================================================
// 15. Array.from() — on Set
// ===================================================================

#[test]
fn array_from_on_set() {
    assert_eq!(
        run_js(
            r#"
const s = new Set([1, 2, 2, 3, 3, 3]);
const arr = Array.from(s);
arr.sort((a, b) => a - b);
console.log(arr.join(","));
"#
        ),
        vec!["1,2,3"]
    );
}

// ===================================================================
// 16. Array.from() — on Map
// ===================================================================

#[test]
fn array_from_on_map() {
    assert_eq!(
        run_js(
            r#"
const m = new Map([["a", 1], ["b", 2], ["c", 3]]);
const arr = Array.from(m);
console.log(arr.length);
console.log(arr[0][0] + ":" + arr[0][1]);
"#
        ),
        vec!["3", "a:1"]
    );
}

// ===================================================================
// 17. Array.of() — creates array from arguments
// ===================================================================

#[test]
fn array_of_from_arguments() {
    assert_eq!(
        run_js(
            r#"
const arr = Array.of(7, 8, 9);
console.log(arr.length);
console.log(arr.join(","));
"#
        ),
        vec!["3", "7,8,9"]
    );
}

// ===================================================================
// 18. groupBy concept via reduce
// ===================================================================

#[test]
fn array_groupby_via_reduce() {
    assert_eq!(
        run_js(
            r#"
const items = ["apple", "avocado", "banana", "blueberry", "cherry"];
const grouped = items.reduce((acc, item) => {
    const key = item[0];
    if (!acc[key]) acc[key] = [];
    acc[key].push(item);
    return acc;
}, {});
console.log(grouped["a"].length);
console.log(grouped["b"].length);
console.log(grouped["c"].length);
"#
        ),
        vec!["2", "2", "1"]
    );
}

// ===================================================================
// 19. Array.prototype.fill() — with start and end indices
// ===================================================================

#[test]
fn array_fill_with_start_end() {
    assert_eq!(
        run_js(
            r#"
const arr = [0, 0, 0, 0, 0];
arr.fill(7, 1, 4);
console.log(arr.join(","));
"#
        ),
        vec!["0,7,7,7,0"]
    );
}

// ===================================================================
// 20. Array.prototype.copyWithin() — copies elements within array
// ===================================================================

#[test]
fn array_copywithin_basic() {
    assert_eq!(
        run_js(
            r#"
const arr = [1, 2, 3, 4, 5];
arr.copyWithin(1, 3, 5);
console.log(arr.join(","));
"#
        ),
        vec!["1,4,5,4,5"]
    );
}

// ===================================================================
// 21. Array.prototype.keys() iterator
// ===================================================================

#[test]
fn array_keys_iterator() {
    assert_eq!(
        run_js(
            r#"
const arr = ["x", "y", "z"];
const keys = [...arr.keys()];
console.log(keys.join(","));
"#
        ),
        vec!["0,1,2"]
    );
}

// ===================================================================
// 22. Array.prototype.values() iterator
// ===================================================================

#[test]
fn array_values_iterator() {
    assert_eq!(
        run_js(
            r#"
const arr = [100, 200, 300];
const vals = [...arr.values()];
console.log(vals.join(","));
"#
        ),
        vec!["100,200,300"]
    );
}

// ===================================================================
// 23. Array.prototype.entries() iterator
// ===================================================================

#[test]
fn array_entries_iterator() {
    assert_eq!(
        run_js(
            r#"
const arr = ["a", "b", "c"];
const pairs = [];
for (const [i, v] of arr.entries()) {
    pairs.push(i + "=" + v);
}
console.log(pairs.join(","));
"#
        ),
        vec!["0=a,1=b,2=c"]
    );
}

// ===================================================================
// 24. Array.prototype.indexOf() — with fromIndex
// ===================================================================

#[test]
fn array_indexof_with_fromindex() {
    assert_eq!(
        run_js(
            r#"
const arr = [1, 2, 3, 2, 1];
console.log(arr.indexOf(2));
console.log(arr.indexOf(2, 2));
console.log(arr.indexOf(2, 4));
"#
        ),
        vec!["1", "3", "-1"]
    );
}

// ===================================================================
// 25. Array.prototype.lastIndexOf() — with fromIndex
// ===================================================================

#[test]
fn array_lastindexof_with_fromindex() {
    assert_eq!(
        run_js(
            r#"
const arr = [1, 2, 3, 2, 1];
console.log(arr.lastIndexOf(2));
console.log(arr.lastIndexOf(2, 2));
console.log(arr.lastIndexOf(9));
"#
        ),
        vec!["3", "1", "-1"]
    );
}

// ===================================================================
// 26. Chaining toSorted + map + filter
// ===================================================================

#[test]
fn array_tosorted_map_filter_chain() {
    assert_eq!(
        run_js(
            r#"
const nums = [5, 2, 8, 1, 9, 3];
const result = nums
    .toSorted((a, b) => a - b)
    .filter(x => x > 3)
    .map(x => x * 10);
console.log(result.join(","));
"#
        ),
        vec!["50,80,90"]
    );
}

// ===================================================================
// 27. Array spread creates shallow copy
// ===================================================================

#[test]
fn array_spread_shallow_copy() {
    assert_eq!(
        run_js(
            r#"
const orig = [1, 2, 3];
const copy = [...orig];
copy.push(4);
console.log(orig.join(","));
console.log(copy.join(","));
console.log(orig === copy);
"#
        ),
        vec!["1,2,3", "1,2,3,4", "false"]
    );
}

// ===================================================================
// 28. Nested array — at() method
// ===================================================================

#[test]
fn nested_array_at_method() {
    assert_eq!(
        run_js(
            r#"
const matrix = [[1, 2], [3, 4], [5, 6]];
console.log(matrix.at(0).at(-1));
console.log(matrix.at(-1).at(0));
"#
        ),
        vec!["2", "5"]
    );
}

// ===================================================================
// 29. Array.prototype.findIndex() — returns -1 when not found
// ===================================================================

#[test]
fn array_findindex_returns_negative_one_when_not_found() {
    assert_eq!(
        run_js(
            r#"
const arr = [1, 2, 3, 4, 5];
console.log(arr.findIndex(x => x > 100));
console.log(arr.findIndex(x => x === 0));
"#
        ),
        vec!["-1", "-1"]
    );
}

// ===================================================================
// 30. Array.prototype.flatMap() — with index parameter
// ===================================================================

#[test]
fn array_flatmap_with_index_parameter() {
    assert_eq!(
        run_js(
            r#"
const arr = ["a", "b", "c"];
const result = arr.flatMap((val, idx) => [idx, val]);
console.log(result.join(","));
"#
        ),
        vec!["0,a,1,b,2,c"]
    );
}
