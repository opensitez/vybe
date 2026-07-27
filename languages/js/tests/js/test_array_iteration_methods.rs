/// Array reduce/reduceRight, every/some/find/findIndex, flat/flatMap,
/// includes, indexOf, copyWithin, fill, entries/keys/values iterators.
use super::helpers::run_js;

#[test]
fn reduce_sum() {
    assert_eq!(
        run_js(
            r#"
const sum = [1, 2, 3, 4, 5].reduce((acc, x) => acc + x, 0);
console.log(sum);
"#
        ),
        vec!["15"]
    );
}

#[test]
fn reduce_no_initial_uses_first_element() {
    assert_eq!(
        run_js(
            r#"
const max = [3, 1, 4, 1, 5, 9].reduce((a, b) => a > b ? a : b);
console.log(max);
"#
        ),
        vec!["9"]
    );
}

#[test]
fn reduce_empty_array_without_initial_throws_type_error() {
    assert_eq!(
        run_js(
            r#"
try {
    [].reduce((a, b) => a + b);
} catch (e) {
    console.log(e.name);
}
"#
        ),
        vec!["TypeError"]
    );
}

#[test]
fn reduceright_processes_right_to_left() {
    assert_eq!(
        run_js(
            r#"
const result = [[1,2],[3,4],[5,6]].reduceRight((acc, x) => acc.concat(x), []);
console.log(result.join(","));
"#
        ),
        vec!["5,6,3,4,1,2"]
    );
}

#[test]
fn reduce_right_with_initial_on_empty_array() {
    assert_eq!(
        run_js(
            r#"
console.log([].reduceRight((a, b) => a + b, 10));
"#
        ),
        vec!["10"]
    );
}

#[test]
fn every_all_satisfy() {
    assert_eq!(
        run_js(
            r#"
console.log([2, 4, 6, 8].every(n => n % 2 === 0));
console.log([2, 3, 6].every(n => n % 2 === 0));
"#
        ),
        vec!["true", "false"]
    );
}

#[test]
fn every_empty_is_vacuously_true() {
    assert_eq!(
        run_js(
            r#"
console.log([].every(() => false));
"#
        ),
        vec!["true"]
    );
}

#[test]
fn some_at_least_one() {
    assert_eq!(
        run_js(
            r#"
console.log([1, 3, 5, 6].some(n => n % 2 === 0));
console.log([1, 3, 5].some(n => n % 2 === 0));
"#
        ),
        vec!["true", "false"]
    );
}

#[test]
fn find_returns_first_match() {
    assert_eq!(
        run_js(
            r#"
const result = [1, 2, 3, 4].find(n => n > 2);
console.log(result);
"#
        ),
        vec!["3"]
    );
}

#[test]
fn find_returns_undefined_when_none() {
    assert_eq!(
        run_js(
            r#"
console.log([1, 2, 3].find(n => n > 10));
"#
        ),
        vec!["undefined"]
    );
}

#[test]
fn find_skips_empty_slots() {
    assert_eq!(
        run_js(
            r#"
const arr = [, 1, , 3];
let seen = [];
const value = arr.find((value, index) => {
    seen.push(index);
    return value === 3;
});
console.log(value);
console.log(seen.join(","));
"#
    ),
        vec!["3", "0,1,2,3"]
    );
}

#[test]
fn findindex_returns_first_matching_index() {
    assert_eq!(
        run_js(
            r#"
console.log([1, 2, 3, 4].findIndex(n => n > 2));
console.log([1, 2, 3].findIndex(n => n > 10));
"#
        ),
        vec!["2", "-1"]
    );
}

#[test]
fn includes_basic() {
    assert_eq!(
        run_js(
            r#"
const arr = [1, 2, NaN, 3];
console.log(arr.includes(2));
console.log(arr.includes(NaN)); // includes uses SameValueZero
console.log(arr.includes(5));
"#
        ),
        vec!["true", "true", "false"]
    );
}

#[test]
fn includes_with_from_index() {
    assert_eq!(
        run_js(
            r#"
const arr = [1, 2, 3, 2, 1];
console.log(arr.includes(2, 2));
console.log(arr.includes(1, -1));
"#
        ),
        vec!["true", "true"]
    );
}

#[test]
fn indexof_uses_strict_equality() {
    assert_eq!(
        run_js(
            r#"
const arr = [1, 2, NaN, 3];
console.log(arr.indexOf(NaN)); // -1 — NaN !== NaN
console.log(arr.indexOf(2));
"#
        ),
        vec!["-1", "1"]
    );
}

#[test]
fn copywithin_basic() {
    assert_eq!(
        run_js(
            r#"
const arr = [1, 2, 3, 4, 5];
arr.copyWithin(0, 3); // copy from index 3 to start
console.log(arr.join(","));
"#
        ),
        vec!["4,5,3,4,5"]
    );
}

#[test]
fn copywithin_with_end() {
    assert_eq!(
        run_js(
            r#"
const arr = [1, 2, 3, 4, 5];
arr.copyWithin(1, 3, 4); // copy arr[3] to arr[1]
console.log(arr.join(","));
"#
        ),
        vec!["1,4,3,4,5"]
    );
}

#[test]
fn fill_entire_array() {
    assert_eq!(
        run_js(
            r#"
const arr = new Array(4).fill(0);
console.log(arr.join(","));
"#
        ),
        vec!["0,0,0,0"]
    );
}

#[test]
fn fill_partial_range() {
    assert_eq!(
        run_js(
            r#"
const arr = [1, 2, 3, 4, 5];
arr.fill(99, 1, 3);
console.log(arr.join(","));
"#
        ),
        vec!["1,99,99,4,5"]
    );
}

#[test]
fn entries_yields_index_value_pairs() {
    assert_eq!(
        run_js(
            r#"
const pairs = [...["a", "b", "c"].entries()];
console.log(pairs.map(([i, v]) => i + ":" + v).join(","));
"#
        ),
        vec!["0:a,1:b,2:c"]
    );
}

#[test]
fn keys_yields_indices() {
    assert_eq!(
        run_js(
            r#"
const keys = [...["a", "b", "c"].keys()];
console.log(keys.join(","));
"#
        ),
        vec!["0,1,2"]
    );
}

#[test]
fn values_yields_elements() {
    assert_eq!(
        run_js(
            r#"
const vals = [...[10, 20, 30].values()];
console.log(vals.join(","));
"#
        ),
        vec!["10,20,30"]
    );
}
