//! Illegal `[]` offset access and `foreach` on non-array types (PHP 8+ TypeErrors).

crate::php_cases! {
    read_offset_on_null_throws_type_error => {
        r#"<?php
$x = null;
try { echo $x[0]; }
catch (TypeError $e) { echo 'null-read'; }
"#,
        ["null-read"]
    };

    write_offset_on_null_throws_type_error => {
        r#"<?php
$x = null;
try { $x[0] = 1; echo 'ok'; }
catch (TypeError $e) { echo 'null-write'; }
"#,
        ["null-write"]
    };

    append_offset_on_null_throws_type_error => {
        r#"<?php
$x = null;
try { $x[] = 1; echo 'ok'; }
catch (TypeError $e) { echo 'null-append'; }
"#,
        ["null-append"]
    };

    read_offset_on_int_throws_type_error => {
        r#"<?php
try { echo 42[0]; }
catch (TypeError $e) { echo 'int-read'; }
"#,
        ["int-read"]
    };

    write_offset_on_int_throws_type_error => {
        r#"<?php
$x = 1;
try { $x[0] = 9; echo 'ok'; }
catch (TypeError $e) { echo 'int-write'; }
"#,
        ["int-write"]
    };

    read_offset_on_float_throws_type_error => {
        r#"<?php
try { echo 1.5[0]; }
catch (TypeError $e) { echo 'float-read'; }
"#,
        ["float-read"]
    };

    read_offset_on_false_throws_type_error => {
        r#"<?php
$x = false;
try { echo $x[0]; }
catch (TypeError $e) { echo 'false-read'; }
"#,
        ["false-read"]
    };

    foreach_on_int_throws_type_error => {
        r#"<?php
try { foreach (1 as $v) { echo $v; } echo 'ok'; }
catch (TypeError $e) { echo 'foreach-int'; }
"#,
        ["foreach-int"]
    };

    foreach_on_null_throws_type_error => {
        r#"<?php
try { foreach (null as $v) { echo $v; } echo 'ok'; }
catch (TypeError $e) { echo 'foreach-null'; }
"#,
        ["foreach-null"]
    };

    unset_offset_on_scalar_throws_type_error => {
        r#"<?php
$x = 5;
try { unset($x[0]); echo 'ok'; }
catch (TypeError $e) { echo 'unset-scalar'; }
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

    array_pop_on_int_throws_type_error => {
        r#"<?php
try { array_pop(1); echo 'ok'; }
catch (TypeError $e) { echo 'pop-int'; }
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

    count_on_non_countable_string_throws_in_strict_shape => {
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

    compact_requires_string_var_names => {
        r#"<?php
try { compact(1); echo 'ok'; }
catch (TypeError $e) { echo 'compact'; }
"#,
        ["compact"]
    };

    increment_offset_on_int_throws => {
        r#"<?php
$x = 1;
try { $x[0]++; echo 'ok'; }
catch (TypeError $e) { echo 'inc-scalar'; }
"#,
        ["inc-scalar"]
    };
}
