crate::js_cases! {
    array_constructor_without_arguments_is_empty => {
        r#"
const arr = new Array();
console.log(arr.length);
console.log(Array.isArray(arr));
"#,
        ["0", "true"]
    };

    array_constructor_with_single_length_creates_holes => {
        r#"
const arr = new Array(3);
console.log(arr.length);
console.log(0 in arr);
console.log(arr[0] === undefined);
"#,
        ["3", "false", "true"]
    };

    array_constructor_with_multiple_arguments_creates_elements => {
        r#"
const arr = new Array(1, 2, 3);
console.log(arr.length);
console.log(arr.join(","));
"#,
        ["3", "1,2,3"]
    };

    array_constructor_with_string_single_argument_creates_element => {
        r#"
const arr = new Array("3");
console.log(arr.length);
console.log(arr[0]);
"#,
        ["1", "3"]
    };

    array_constructor_with_fractional_length_throws_range_error => {
        r#"
try {
  new Array(2.5);
  console.log("no error");
} catch (error) {
  console.log(error instanceof RangeError);
}
"#,
        ["true"]
    };

    array_constructor_with_negative_length_throws_range_error => {
        r#"
try {
  new Array(-1);
  console.log("no error");
} catch (error) {
  console.log(error instanceof RangeError);
}
"#,
        ["true"]
    };

    array_length_truncation_removes_tail_elements => {
        r#"
const arr = [1, 2, 3, 4];
arr.length = 2;
console.log(arr.join(","));
console.log(2 in arr);
"#,
        ["1,2", "false"]
    };

    array_length_extension_creates_new_holes => {
        r#"
const arr = [1, 2];
arr.length = 4;
console.log(arr.length);
console.log(2 in arr);
console.log(arr[3] === undefined);
"#,
        ["4", "false", "true"]
    };

    array_length_zero_clears_all_indices => {
        r#"
const arr = [1, 2, 3];
arr.length = 0;
console.log(arr.length);
console.log(0 in arr);
"#,
        ["0", "false"]
    };

    sparse_array_join_emits_empty_segments_for_holes => {
        r#"
const arr = [];
arr[1] = "x";
arr[3] = "y";
console.log(arr.join(","));
"#,
        [",x,,y"]
    };

    sparse_array_object_keys_list_only_present_indices => {
        r#"
const arr = [];
arr[1] = "x";
arr[3] = "y";
console.log(Object.keys(arr).join(","));
"#,
        ["1,3"]
    };

    sparse_array_reading_hole_returns_undefined => {
        r#"
const arr = [];
arr[2] = "x";
console.log(arr[1] === undefined);
"#,
        ["true"]
    };

    sparse_array_in_operator_distinguishes_hole_from_present_undefined => {
        r#"
const arr = [undefined, , undefined];
console.log(0 in arr);
console.log(1 in arr);
console.log(2 in arr);
"#,
        ["true", "false", "true"]
    };

    sparse_array_for_each_skips_holes => {
        r#"
const arr = [, "a", , "b"];
const seen = [];
arr.forEach((value, index) => seen.push(index + ":" + value));
console.log(seen.join(","));
"#,
        ["1:a,3:b"]
    };

    sparse_array_map_preserves_hole_positions => {
        r#"
const arr = [, 2, , 4];
const mapped = arr.map(x => x * 2);
console.log(mapped.length);
console.log(0 in mapped);
console.log(1 in mapped);
console.log(2 in mapped);
console.log(3 in mapped);
console.log(mapped[1]);
"#,
        ["4", "false", "true", "false", "true", "4"]
    };

    sparse_array_filter_skips_holes_and_compacts_result => {
        r#"
const arr = [, 1, , 2];
const filtered = arr.filter(() => true);
console.log(filtered.length);
console.log(filtered.join(","));
"#,
        ["2", "1,2"]
    };

    sparse_array_some_checks_only_present_elements => {
        r#"
const arr = [, , 3];
const seen = [];
const result = arr.some((value, index) => {
  seen.push(index);
  return value === 3;
});
console.log(result);
console.log(seen.join(","));
"#,
        ["true", "2"]
    };

    sparse_array_every_ignores_holes_for_predicate_calls => {
        r#"
const arr = [, , 4];
const seen = [];
const result = arr.every((value, index) => {
  seen.push(index);
  return value > 0;
});
console.log(result);
console.log(seen.join(","));
"#,
        ["true", "2"]
    };

    sparse_array_reduce_uses_first_present_element_without_initial => {
        r#"
const arr = [];
arr[2] = 5;
arr[4] = 7;
console.log(arr.reduce((acc, value) => acc + value));
"#,
        ["12"]
    };

    sparse_array_reduce_right_walks_present_elements_from_end => {
        r#"
const arr = [];
arr[1] = "b";
arr[3] = "d";
arr[5] = "f";
console.log(arr.reduceRight((acc, value) => acc + value, ""));
"#,
        ["fdb"]
    };

    sparse_array_includes_treats_hole_as_undefined => {
        r#"
console.log([,].includes(undefined));
"#,
        ["true"]
    };

    sparse_array_indexof_does_not_match_hole_as_undefined => {
        r#"
console.log([,].indexOf(undefined));
"#,
        ["-1"]
    };

    delete_array_element_preserves_length_and_creates_hole => {
        r#"
const arr = [1, 2, 3];
delete arr[1];
console.log(arr.length);
console.log(1 in arr);
"#,
        ["3", "false"]
    };

    pop_on_trailing_hole_returns_undefined_and_shrinks_length => {
        r#"
const arr = [1, 2, 3];
arr.length = 4;
console.log(arr.pop() === undefined);
console.log(arr.length);
console.log(arr.join(","));
"#,
        ["true", "3", "1,2,3"]
    };

    push_after_length_extension_appends_at_new_end => {
        r#"
const arr = [1];
arr.length = 3;
arr.push(4);
console.log(arr.length);
console.log(arr[3]);
console.log(Object.keys(arr).join(","));
"#,
        ["4", "4", "0,3"]
    };

    shift_on_sparse_array_moves_present_elements_left => {
        r#"
const arr = [];
arr[1] = "x";
arr[3] = "y";
const first = arr.shift();
console.log(first === undefined);
console.log(arr.length);
console.log(Object.keys(arr).join(","));
"#,
        ["true", "3", "0,2"]
    };

    unshift_on_sparse_array_preserves_holes_between_shifted_indices => {
        r#"
const arr = [];
arr[1] = "x";
arr[2] = "y";
arr.unshift("a");
console.log(arr.length);
console.log(arr.join(","));
console.log(Object.keys(arr).join(","));
"#,
        ["4", "a,,x,y", "0,2,3"]
    };

    sparse_array_flat_removes_holes => {
        r#"
const arr = [1, , 3, , 5];
console.log(arr.flat().length + "|" + arr.flat().join(","));
"#,
        ["3|1,3,5"]
    };
}

