//! Built-in functions rejecting wrong argument types or shapes at runtime.

crate::php_cases! {
    strlen_on_array_throws_type_error => {
        r#"<?php
try { strlen([1, 2]); echo 'ok'; }
catch (TypeError $e) { echo 'strlen-arr'; }
"#,
        ["strlen-arr"]
    };

    strlen_on_int_throws_type_error => {
        r#"<?php
try { strlen(99); echo 'ok'; }
catch (TypeError $e) { echo 'strlen-int'; }
"#,
        ["strlen-int"]
    };

    array_merge_on_string_throws_type_error => {
        r#"<?php
try { array_merge('ab', [1]); echo 'ok'; }
catch (TypeError $e) { echo 'merge-str'; }
"#,
        ["merge-str"]
    };

    array_combine_length_mismatch_returns_false => {
        r#"<?php
$ok = array_combine(['a', 'b'], [1]);
echo $ok === false ? 'false' : 'map';
"#,
        ["false"]
    };

    array_combine_empty_arrays_returns_empty => {
        r#"<?php
$m = array_combine([], []);
echo is_array($m) && count($m) === 0 ? 'empty' : 'other';
"#,
        ["empty"]
    };

    array_map_callback_not_callable_throws => {
        r#"<?php
try { array_map('not-a-function', [1]); echo 'ok'; }
catch (TypeError $e) { echo 'map-cb'; }
"#,
        ["map-cb"]
    };

    array_filter_on_object_without_traversable_throws => {
        r#"<?php
try { array_filter(new stdClass()); echo 'ok'; }
catch (TypeError $e) { echo 'filter-obj'; }
"#,
        ["filter-obj"]
    };

    implode_glue_must_be_string => {
        r#"<?php
try { implode([1, 2], '-'); echo 'ok'; }
catch (TypeError $e) { echo 'implode'; }
"#,
        ["implode"]
    };

    explode_separator_must_be_string => {
        r#"<?php
try { explode(1, 'a,b'); echo 'ok'; }
catch (TypeError $e) { echo 'explode'; }
"#,
        ["explode"]
    };

    str_contains_haystack_must_be_string => {
        r#"<?php
try { str_contains(1, 'x'); echo 'ok'; }
catch (TypeError $e) { echo 'contains'; }
"#,
        ["contains"]
    };

    str_starts_with_subject_must_be_string => {
        r#"<?php
try { str_starts_with([], 'a'); echo 'ok'; }
catch (TypeError $e) { echo 'starts'; }
"#,
        ["starts"]
    };

    max_on_empty_array_returns_null_in_php8 => {
        r#"<?php
$v = max([]);
echo $v === null ? 'null' : (string)$v;
"#,
        ["null"]
    };

    min_on_empty_array_returns_null_in_php8 => {
        r#"<?php
$v = min([]);
echo $v === null ? 'null' : (string)$v;
"#,
        ["null"]
    };

    abs_requires_numeric => {
        r#"<?php
try { abs([]); echo 'ok'; }
catch (TypeError $e) { echo 'abs-arr'; }
"#,
        ["abs-arr"]
    };

    round_on_array_throws_type_error => {
        r#"<?php
try { round([1.2]); echo 'ok'; }
catch (TypeError $e) { echo 'round-arr'; }
"#,
        ["round-arr"]
    };

    sort_on_non_array_throws_type_error => {
        r#"<?php
try { sort('abc'); echo 'ok'; }
catch (TypeError $e) { echo 'sort-str'; }
"#,
        ["sort-str"]
    };

    ksort_on_object_throws_type_error => {
        r#"<?php
try { ksort(new stdClass()); echo 'ok'; }
catch (TypeError $e) { echo 'ksort-obj'; }
"#,
        ["ksort-obj"]
    };

    array_slice_on_string_throws_type_error => {
        r#"<?php
try { array_slice('abc', 0); echo 'ok'; }
catch (TypeError $e) { echo 'slice-str'; }
"#,
        ["slice-str"]
    };

    array_keys_on_int_throws_type_error => {
        r#"<?php
try { array_keys(1); echo 'ok'; }
catch (TypeError $e) { echo 'keys-int'; }
"#,
        ["keys-int"]
    };

    array_values_on_bool_throws_type_error => {
        r#"<?php
try { array_values(true); echo 'ok'; }
catch (TypeError $e) { echo 'values-bool'; }
"#,
        ["values-bool"]
    };

    preg_match_pattern_must_be_string => {
        r#"<?php
try { preg_match(1, 'hay'); echo 'ok'; }
catch (TypeError $e) { echo 'preg'; }
"#,
        ["preg"]
    };
}
