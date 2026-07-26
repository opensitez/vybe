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

    isset_offset_on_array_key => {
        r#"<?php
$x = [0 => 'a', 1 => null];
echo (isset($x[0]) ? 'a' : 'na') . '|';
echo (isset($x[1]) ? 'b' : 'nb') . '|';
echo (isset($x[2]) ? 'c' : 'nc');
"#,
        ["a|b|nc"]
    };

    empty_offset_on_array_key => {
        r#"<?php
$x = ['k' => 0];
echo (empty($x['k']) ? 'empty' : 'not');
"#,
        ["empty"]
    };

    isset_offset_on_scalar_returns_false => {
        r#"<?php
$x = 42;
echo (isset($x[0]) ? 'set' : 'unset') . '|';
echo (isset($x['a']) ? 'set' : 'unset');
"#,
        ["unset|unset"]
    };

    unset_offset_on_array_removes_key => {
        r#"<?php
$x = ['a' => 1, 'b' => 2];
unset($x['a']);
echo (array_key_exists('a', $x) ? 'yes' : 'no');
"#,
        ["no"]
    };

    read_negative_offset_on_string => {
        r#"<?php
echo 'xyz'[-2];
"#,
        ["y"]
    };

    write_string_offset_mutates_byte => {
        r#"<?php
$s = 'abc';
$s[1] = 'Z';
echo $s;
"#,
        ["aZc"]
    };

    append_string_offset_out_of_range_raises => {
        r#"<?php
$s = 'ab';
try { $s[] = 'c'; echo $s; }
catch (TypeError $e) { echo 'err'; }
catch (\Error $e) { echo 'err'; }
"#,
        ["err"]
    };

    read_missing_array_offset_returns_null => {
        r#"<?php
$x = ['a' => 1];
echo array_key_exists('b', $x) ? 'yes' : 'no';
echo '|';
$v = $x['b'];
echo $v === null ? 'null' : 'set';
"#,
        ["no|null"]
    };

    isset_missing_numeric_offset_in_array => {
        r#"<?php
$x = [1, 2, 3];
echo isset($x[4]) ? 'yes' : 'no';
"#,
        ["no"]
    };

    read_array_offset_on_truthy_string_key => {
        r#"<?php
$x = ['10' => 'ten', 'foo' => 'bar'];
echo $x['10'];
echo '|';
echo $x[10];
"#,
        ["ten|ten"]
    };

    unset_string_offset_throws => {
        r#"<?php
$x = 'hello';
try { unset($x[1]); echo 'ok'; }
catch (\Error $e) { echo 'unset-string'; }
catch (TypeError $e) { echo 'unset-string'; }
"#,
        ["unset-string"]
    };

    string_offset_out_of_range_read_returns_null => {
        r#"<?php
$x = 'hi';
echo $x[5] === null ? 'null' : 'set';
"#,
        ["null"]
    };

    write_string_offset_with_non_scalar_throws => {
        r#"<?php
$x = 'ab';
try { $x[0] = [1]; echo 'ok'; }
catch (TypeError $e) { echo 'type'; }
catch (\Error $e) { echo 'type'; }
"#,
        ["type"]
    };

    foreach_on_array_with_numeric_string_keys => {
        r#"<?php
$x = ['0' => 'zero', 1 => 'one', '2' => 'two'];
$out = [];
foreach ($x as $v) { $out[] = $v; }
echo implode('|', $out);
"#,
        ["zero|one|two"]
    };
}
