//! Built-in functions rejecting wrong argument types or shapes at runtime.
//!
//! Verified against the `php` 8.4 CLI. Notes on corrected expectations:
//!   - array/object → string/number is a `TypeError` in every mode;
//!   - scalar → string IS allowed (weak coercion): `strlen(99)`, `explode(1,..)`,
//!     `str_contains(1,..)`, `preg_match(1,..)` do NOT throw;
//!   - `array_combine` length-mismatch and `max([])`/`min([])` throw
//!     `ValueError` in PHP 8 (they returned `false`/`null` in PHP 7);
//!   - `sort('literal')` throws base `\Error` (by-ref of a non-variable),
//!     not `TypeError`.

crate::php_cases! {
    strlen_on_array_throws_type_error => {
        r#"<?php
try { strlen([1, 2]); echo 'ok'; }
catch (TypeError $e) { echo 'strlen-arr'; }
"#,
        ["strlen-arr"]
    };

    strlen_on_int_coerces_and_returns_length => {
        r#"<?php
echo strlen(99);
"#,
        ["2"]
    };

    array_merge_on_string_throws_type_error => {
        r#"<?php
try { array_merge('ab', [1]); echo 'ok'; }
catch (TypeError $e) { echo 'merge-str'; }
"#,
        ["merge-str"]
    };

    array_combine_length_mismatch_throws_value_error => {
        r#"<?php
try { array_combine(['a', 'b'], [1]); echo 'ok'; }
catch (ValueError $e) { echo 'mismatch'; }
"#,
        ["mismatch"]
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

    implode_array_glue_swapped_still_joins => {
        r#"<?php
try { implode([1, 2], '-'); echo 'ok'; }
catch (TypeError $e) { echo 'implode'; }
"#,
        ["implode"]
    };

    explode_separator_coerces_scalar => {
        r#"<?php
try { explode(1, 'a,b'); echo 'ok'; }
catch (TypeError $e) { echo 'explode'; }
"#,
        ["ok"]
    };

    str_contains_haystack_coerces_scalar => {
        r#"<?php
try { str_contains(1, 'x'); echo 'ok'; }
catch (TypeError $e) { echo 'contains'; }
"#,
        ["ok"]
    };

    str_starts_with_subject_must_be_string => {
        r#"<?php
try { str_starts_with([], 'a'); echo 'ok'; }
catch (TypeError $e) { echo 'starts'; }
"#,
        ["starts"]
    };

    max_on_empty_array_throws_value_error => {
        r#"<?php
try { max([]); echo 'ok'; }
catch (ValueError $e) { echo 'max-empty'; }
"#,
        ["max-empty"]
    };

    min_on_empty_array_throws_value_error => {
        r#"<?php
try { min([]); echo 'ok'; }
catch (ValueError $e) { echo 'min-empty'; }
"#,
        ["min-empty"]
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

    sort_on_literal_throws_error => {
        r#"<?php
try { sort('abc'); echo 'ok'; }
catch (\Error $e) { echo 'sort-str'; }
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

    preg_match_pattern_coerces_scalar => {
        r#"<?php
try { preg_match(1, 'hay'); echo 'ok'; }
catch (TypeError $e) { echo 'preg'; }
"#,
        ["ok"]
    };
}
