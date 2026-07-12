//! `[]` offset access and `foreach` on non-array types — real PHP 8.4 semantics.
//!
//! Verified against the `php` 8.4 CLI (used purely as a compliance oracle):
//!   - scalar-offset READS warn and yield `null` (NOT a throw);
//!   - `null` auto-vivifies to an array on WRITE/append (no throw);
//!   - scalar-offset WRITE / `unset` / increment throw base `\Error`
//!     ("Cannot use a scalar value as an array"), NOT `TypeError`;
//!   - `foreach` on a non-iterable warns and skips the loop (no throw);
//!   - builtins that reject wrong argument types throw `TypeError`.
//! Earlier revisions of this file asserted `TypeError` everywhere, which real
//! PHP does not do — those expectations were corrected to match the CLI.

crate::php_cases! {
    read_offset_on_null_yields_null => {
        r#"<?php
$x = null;
$v = $x[0];
echo $v === null ? 'null' : 'set';
"#,
        ["null"]
    };

    write_offset_on_null_does_not_throw => {
        r#"<?php
$x = null;
try { $x[0] = 1; echo 'ok'; }
catch (\Throwable $e) { echo 'threw'; }
"#,
        ["ok"]
    };

    append_offset_on_null_does_not_throw => {
        r#"<?php
$x = null;
try { $x[] = 7; echo 'ok'; }
catch (\Throwable $e) { echo 'threw'; }
"#,
        ["ok"]
    };

    read_offset_on_int_yields_null => {
        r#"<?php
$x = 42;
$v = $x[0];
echo $v === null ? 'null' : 'set';
"#,
        ["null"]
    };

    write_offset_on_scalar_throws_error => {
        r#"<?php
$x = 1;
try { $x[0] = 9; echo 'ok'; }
catch (\Error $e) { echo 'int-write'; }
"#,
        ["int-write"]
    };

    read_offset_on_float_yields_null => {
        r#"<?php
$x = 1.5;
$v = $x[0];
echo $v === null ? 'null' : 'set';
"#,
        ["null"]
    };

    read_offset_on_false_yields_null => {
        r#"<?php
$x = false;
$v = $x[0];
echo $v === null ? 'null' : 'set';
"#,
        ["null"]
    };

    foreach_on_int_warns_and_skips => {
        r#"<?php
foreach (1 as $v) { echo $v; }
echo 'done';
"#,
        ["done"]
    };

    foreach_on_null_warns_and_skips => {
        r#"<?php
foreach (null as $v) { echo $v; }
echo 'done';
"#,
        ["done"]
    };

    unset_offset_on_scalar_throws_error => {
        r#"<?php
$x = 5;
try { unset($x[0]); echo 'ok'; }
catch (\Error $e) { echo 'unset-scalar'; }
"#,
        ["unset-scalar"]
    };

    string_byte_offset_read_still_works => {
        r#"<?php
echo 'abc'[1];
"#,
        ["b"]
    };

    string_negative_byte_offset_reads_from_end => {
        r#"<?php
echo 'php'[-1];
"#,
        ["p"]
    };

    array_push_on_string_throws_type_error => {
        r#"<?php
$s = 'hi';
try { array_push($s, 1); echo 'ok'; }
catch (TypeError $e) { echo 'push-str'; }
"#,
        ["push-str"]
    };

    array_pop_on_int_literal_throws_error => {
        r#"<?php
try { array_pop(1); echo 'ok'; }
catch (\Error $e) { echo 'pop-int'; }
"#,
        ["pop-int"]
    };

    array_key_exists_on_int_throws_type_error => {
        r#"<?php
try { array_key_exists(0, 1); echo 'ok'; }
catch (TypeError $e) { echo 'key-int'; }
"#,
        ["key-int"]
    };

    in_array_with_non_array_haystack_throws => {
        r#"<?php
try { in_array(1, 42); echo 'ok'; }
catch (TypeError $e) { echo 'in-array'; }
"#,
        ["in-array"]
    };

    count_on_non_countable_string_throws => {
        r#"<?php
try { count('abc'); echo 'ok'; }
catch (TypeError $e) { echo 'count-str'; }
catch (ValueError $e) { echo 'count-str'; }
"#,
        ["count-str"]
    };

    iterable_to_array_on_int_throws => {
        r#"<?php
try { iterator_to_array(7); echo 'ok'; }
catch (TypeError $e) { echo 'iter-int'; }
"#,
        ["iter-int"]
    };

    compact_non_string_name_does_not_throw => {
        r#"<?php
try { compact(1); echo 'ok'; }
catch (\Throwable $e) { echo 'threw'; }
"#,
        ["ok"]
    };

    increment_offset_on_scalar_throws_error => {
        r#"<?php
$x = 1;
try { $x[0]++; echo 'ok'; }
catch (\Error $e) { echo 'inc-scalar'; }
"#,
        ["inc-scalar"]
    };
}
