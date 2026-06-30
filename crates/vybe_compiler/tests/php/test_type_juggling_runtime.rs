//! Loose vs strict comparison and PHP type juggling (runtime).

crate::php_cases! {
    loose_equal_string_one_and_int_one => {
        r#"<?php
echo '1' == 1 ? 'eq' : 'ne';
"#,
        ["eq"]
    };

    strict_not_equal_string_one_and_int_one => {
        r#"<?php
echo '1' === 1 ? 'eq' : 'ne';
"#,
        ["ne"]
    };

    loose_equal_empty_string_and_zero => {
        r#"<?php
echo '' == 0 ? 'eq' : 'ne';
"#,
        ["eq"]
    };

    strict_empty_string_not_equal_zero => {
        r#"<?php
echo '' === 0 ? 'eq' : 'ne';
"#,
        ["ne"]
    };

    loose_null_equals_false => {
        r#"<?php
echo null == false ? 'eq' : 'ne';
"#,
        ["eq"]
    };

    strict_null_not_identical_false => {
        r#"<?php
echo null === false ? 'eq' : 'ne';
"#,
        ["ne"]
    };

    loose_array_equals_false => {
        r#"<?php
echo [] == false ? 'eq' : 'ne';
"#,
        ["eq"]
    };

    spaceship_int_less => {
        r#"<?php
echo 1 <=> 2;
"#,
        ["-1"]
    };

    spaceship_int_greater => {
        r#"<?php
echo 3 <=> 2;
"#,
        ["1"]
    };

    spaceship_equal => {
        r#"<?php
echo 4 <=> 4;
"#,
        ["0"]
    };

    string_concatenates_with_int => {
        r#"<?php
echo 'n' . 7;
"#,
        ["n7"]
    };

    int_plus_string_numeric => {
        r#"<?php
echo 2 + '3';
"#,
        ["5"]
    };

    float_string_multiplication => {
        r#"<?php
echo (int)(2.5 * '2');
"#,
        ["5"]
    };

    boolean_plus_boolean => {
        r#"<?php
echo true + true;
"#,
        ["2"]
    };

    null_coalesce_on_undefined => {
        r#"<?php
echo $missing ?? 9;
"#,
        ["9"]
    };

    null_coalesce_chain => {
        r#"<?php
$a = null;
$b = null;
echo $a ?? $b ?? 'z';
"#,
        ["z"]
    };

    empty_string_is_empty => {
        r#"<?php
echo empty('') ? 'yes' : 'no';
"#,
        ["yes"]
    };

    empty_zero_is_empty => {
        r#"<?php
echo empty(0) ? 'yes' : 'no';
"#,
        ["yes"]
    };

    empty_string_zero_not_empty_strict => {
        r#"<?php
echo empty('0') ? 'yes' : 'no';
"#,
        ["no"]
    };

    isset_on_null_value => {
        r#"<?php
$x = null;
echo isset($x) ? 'yes' : 'no';
"#,
        ["no"]
    };

    is_numeric_on_numeric_string => {
        r#"<?php
echo is_numeric('12.3') ? 'yes' : 'no';
"#,
        ["yes"]
    };

    is_numeric_on_hex_string => {
        r#"<?php
echo is_numeric('0x10') ? 'yes' : 'no';
"#,
        ["yes"]
    };

    filter_var_int_validation => {
        r#"<?php
echo filter_var('42', FILTER_VALIDATE_INT);
"#,
        ["42"]
    };

    filter_var_bool_on_string_true => {
        r#"<?php
echo filter_var('true', FILTER_VALIDATE_BOOL) ? '1' : '0';
"#,
        ["1"]
    };

    settype_string_to_int => {
        r#"<?php
$s = '15';
settype($s, 'integer');
echo $s + 1;
"#,
        ["16"]
    };

    settype_int_to_bool_false => {
        r#"<?php
$n = 0;
settype($n, 'boolean');
echo $n ? '1' : '0';
"#,
        ["0"]
    };
}
