/// Array methods not heavily covered — flat, flatMap, at, findLast/findLastIndex,
/// toSorted, toReversed, toSpliced, with (non-mutating), group/groupToMap.
use super::helpers::run_js;

// ── Array.prototype.flat ──────────────────────────────────────────────────────

#[test]
fn flat_depth_one() {
    assert_eq!(
        run_js(
            r#"
const arr = [1, [2, 3], [4, [5]]];
console.log(arr.flat().join(","));
"#
        ),
        vec!["1,2,3,4,5"]
    );
}

#[test]
fn flat_default_depth_is_one() {
    assert_eq!(
        run_js(
            r#"
const arr = [[1, 2], [3, [4, 5]]];
const result = arr.flat();
console.log(result.length);
console.log(Array.isArray(result[2]));
"#
        ),
        vec!["4", "false"]
    );
}

#[test]
fn flat_depth_two() {
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

#[test]
fn flat_infinity_depth() {
    assert_eq!(
        run_js(
            r#"
const deeply = [[[[[1], 2], 3], 4], 5];
console.log(deeply.flat(Infinity).join(","));
"#
        ),
        vec!["1,2,3,4,5"]
    );
}

#[test]
fn flat_removes_empty_holes() {
    assert_eq!(
        run_js(
            r#"
const arr = [1, , 3, [4, , 6]];
const result = arr.flat();
console.log(result.join(","));
"#
        ),
        vec!["1,3,4,6"]
    );
}

// ── Array.prototype.flatMap ───────────────────────────────────────────────────

#[test]
fn flatmap_maps_and_flattens_one_level() {
    assert_eq!(
        run_js(
            r#"
const result = [1, 2, 3].flatMap(x => [x, x * 2]);
console.log(result.join(","));
"#
        ),
        vec!["1,2,2,4,3,6"]
    );
}

#[test]
fn flatmap_can_filter_by_returning_empty() {
    assert_eq!(
        run_js(
            r#"
const result = [1, 2, 3, 4].flatMap(x => x % 2 === 0 ? [x] : []);
console.log(result.join(","));
"#
        ),
        vec!["2,4"]
    );
}

#[test]
fn flatmap_sentence_split() {
    assert_eq!(
        run_js(
            r#"
const sentences = ["Hello World", "Foo Bar"];
const words = sentences.flatMap(s => s.split(" "));
console.log(words.join(","));
"#
        ),
        vec!["Hello,World,Foo,Bar"]
    );
}

// ── Array.prototype.at ────────────────────────────────────────────────────────

#[test]
fn at_positive_index() {
    assert_eq!(
        run_js(
            r#"
const arr = [10, 20, 30, 40];
console.log(arr.at(0));
console.log(arr.at(2));
"#
        ),
        vec!["10", "30"]
    );
}

#[test]
fn at_negative_index() {
    assert_eq!(
        run_js(
            r#"
const arr = [10, 20, 30, 40];
console.log(arr.at(-1));
console.log(arr.at(-2));
"#
        ),
        vec!["40", "30"]
    );
}

#[test]
fn at_out_of_bounds_returns_undefined() {
    assert_eq!(
        run_js(
            r#"
const arr = [1, 2, 3];
console.log(arr.at(10));
console.log(arr.at(-10));
"#
        ),
        vec!["undefined", "undefined"]
    );
}

// ── findLast / findLastIndex ──────────────────────────────────────────────────

#[test]
fn findlast_returns_last_matching() {
    assert_eq!(
        run_js(
            r#"
const arr = [1, 2, 3, 4, 5];
console.log(arr.findLast(x => x % 2 === 0));
"#
        ),
        vec!["4"]
    );
}

#[test]
fn findlast_returns_undefined_if_none() {
    assert_eq!(
        run_js(
            r#"
const arr = [1, 3, 5];
console.log(arr.findLast(x => x % 2 === 0));
"#
        ),
        vec!["undefined"]
    );
}

#[test]
fn findlastindex_returns_last_matching_index() {
    assert_eq!(
        run_js(
            r#"
const arr = [1, 2, 3, 4, 5];
console.log(arr.findLastIndex(x => x % 2 === 0));
"#
        ),
        vec!["3"]
    );
}

#[test]
fn findlastindex_returns_minus_one_if_none() {
    assert_eq!(
        run_js(
            r#"
const arr = [1, 3, 5];
console.log(arr.findLastIndex(x => x % 2 === 0));
"#
        ),
        vec!["-1"]
    );
}

// ── toSorted (non-mutating) ───────────────────────────────────────────────────

#[test]
fn tosorted_returns_new_sorted_array() {
    assert_eq!(
        run_js(
            r#"
const arr = [3, 1, 4, 1, 5, 9];
const sorted = arr.toSorted();
console.log(sorted.join(","));
console.log(arr[0]); // original unchanged
"#
        ),
        vec!["1,1,3,4,5,9", "3"]
    );
}

#[test]
fn tosorted_with_comparator() {
    assert_eq!(
        run_js(
            r#"
const arr = [10, 3, 7, 1];
const sorted = arr.toSorted((a, b) => b - a);
console.log(sorted.join(","));
"#
        ),
        vec!["10,7,3,1"]
    );
}

// ── toReversed (non-mutating) ─────────────────────────────────────────────────

#[test]
fn toreversed_returns_new_reversed_array() {
    assert_eq!(
        run_js(
            r#"
const arr = [1, 2, 3, 4];
const rev = arr.toReversed();
console.log(rev.join(","));
console.log(arr.join(","));
"#
        ),
        vec!["4,3,2,1", "1,2,3,4"]
    );
}

// ── toSpliced (non-mutating) ──────────────────────────────────────────────────

#[test]
fn tospliced_inserts_without_mutating() {
    assert_eq!(
        run_js(
            r#"
const arr = [1, 2, 3, 4];
const spliced = arr.toSpliced(2, 0, 99);
console.log(spliced.join(","));
console.log(arr.join(","));
"#
        ),
        vec!["1,2,99,3,4", "1,2,3,4"]
    );
}

#[test]
fn tospliced_removes_elements() {
    assert_eq!(
        run_js(
            r#"
const arr = [1, 2, 3, 4, 5];
const result = arr.toSpliced(1, 2);
console.log(result.join(","));
"#
        ),
        vec!["1,4,5"]
    );
}

// ── Array.prototype.with (non-mutating) ───────────────────────────────────────

#[test]
fn with_returns_new_array_with_replaced_element() {
    assert_eq!(
        run_js(
            r#"
const arr = [1, 2, 3, 4];
const result = arr.with(2, 99);
console.log(result.join(","));
console.log(arr[2]); // original unchanged
"#
        ),
        vec!["1,2,99,4", "3"]
    );
}

#[test]
fn with_negative_index() {
    assert_eq!(
        run_js(
            r#"
const arr = [1, 2, 3, 4];
const result = arr.with(-1, 99);
console.log(result.join(","));
"#
        ),
        vec!["1,2,3,99"]
    );
}

#[test]
fn with_out_of_bounds_index_throws_rangeerror() {
    assert_eq!(
        run_js(
            r#"
try {
    [1, 2, 3].with(10, 99);
    console.log("no error");
} catch (e) {
    console.log(e instanceof RangeError);
}
"#
        ),
        vec!["true"]
    );
}

