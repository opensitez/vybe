/// Array hole/sparse array behavior — holes vs undefined, methods on holes
use super::helpers::run_js;

#[test]
fn hole_in_array_literal() {
    assert_eq!(
        run_js(
            r#"
const arr = [1,  3]; // hole at index 1
console.log(arr.length);
console.log(arr[1]);       // undefined
console.log(1 in arr);     // false — hole, not undefined
"#
        ),
        vec!["3", "undefined", "false"]
    );
}

#[test]
fn map_skips_holes() {
    assert_eq!(
        run_js(
            r#"
const arr = [1,  3];
const result = arr.map(x => x * 2);
console.log(result[0]);
console.log(1 in result);  // hole preserved in map
console.log(result[2]);
"#
        ),
        vec!["2", "false", "6"]
    );
}

#[test]
fn filter_skips_holes() {
    assert_eq!(
        run_js(
            r#"
const arr = [1,  2,  3];
const result = arr.filter(x => x > 1);
console.log(result.join(","));
"#
        ),
        vec!["2,3"]
    );
}

#[test]
fn foreach_skips_holes() {
    assert_eq!(
        run_js(
            r#"
const arr = [1,  2,  3];
const visited = [];
arr.forEach(x => visited.push(x));
console.log(visited.join(","));
"#
        ),
        vec!["1,2,3"]
    );
}

#[test]
fn reduce_skips_holes() {
    assert_eq!(
        run_js(
            r#"
const arr = [1,  2,  3];
const sum = arr.reduce((acc, x) => acc + x, 0);
console.log(sum);
"#
        ),
        vec!["6"]
    );
}

#[test]
fn join_treats_hole_as_empty() {
    assert_eq!(
        run_js(
            r#"
const arr = [1,  3];
console.log(arr.join(","));
"#
        ),
        vec!["1, 3"]
    );
}

#[test]
fn spread_fills_holes_with_undefined() {
    assert_eq!(
        run_js(
            r#"
const sparse = [1,  3];
const dense = [...sparse];
console.log(1 in dense); // true — undefined, not hole
console.log(dense[1]);
"#
        ),
        vec!["true", "undefined"]
    );
}

#[test]
fn array_from_hole_array_fills_undefined() {
    assert_eq!(
        run_js(
            r#"
const sparse = [1,  3];
const dense = Array.from(sparse);
console.log(1 in dense);
console.log(dense[1]);
"#
        ),
        vec!["true", "undefined"]
    );
}

#[test]
fn find_treats_holes_as_undefined() {
    assert_eq!(
        run_js(
            r#"
const arr = [1,  3];
const found = arr.find(x => x === undefined);
console.log(found);
"#
        ),
        vec!["undefined"]
    );
}

#[test]
fn flat_removes_holes() {
    assert_eq!(
        run_js(
            r#"
const arr = [1,  2,  3];
const flat = arr.flat();
console.log(flat.length);
console.log(flat.join(","));
"#
        ),
        vec!["3", "1,2,3"]
    );
}

#[test]
fn delete_creates_hole() {
    assert_eq!(
        run_js(
            r#"
const arr = [1, 2, 3];
delete arr[1];
console.log(arr.length);  // unchanged
console.log(1 in arr);    // false — hole
console.log(arr[1]);      // undefined
"#
        ),
        vec!["3", "false", "undefined"]
    );
}

#[test]
fn at_and_inclusive_checks() {
    assert_eq!(
        run_js(
            r#"
const arr = [1,  3];
console.log(arr.at(1));
console.log(arr[1]);
console.log(1 in arr);
"#
        ),
        vec!["undefined", "undefined", "false"]
    );
}

#[test]
fn fill_converts_hole_to_value() {
    assert_eq!(
        run_js(
            r#"
const arr = [1,  3];
arr.fill(0, 1, 2);
console.log(arr.length);
console.log(1 in arr);
console.log(arr.join(","));
"#
        ),
        vec!["3", "true", "1,0,3"]
    );
}

#[test]
fn copywithin_preserves_length_and_holes() {
    assert_eq!(
        run_js(
            r#"
const arr = [1,  2,  3];
const copied = arr.copyWithin(1, 0, 2);
console.log(copied.length);
console.log(2 in copied);
console.log(copied.join(","));
"#
        ),
        vec!["5", "true", "1, , 3"]
    );
}

#[test]
fn slice_preserves_sparse_holes() {
    assert_eq!(
        run_js(
            r#"
const arr = [1,  3];
const sliced = arr.slice(0, 3);
console.log(1 in sliced);
"#
        ),
        vec!["false"]
    );
}

