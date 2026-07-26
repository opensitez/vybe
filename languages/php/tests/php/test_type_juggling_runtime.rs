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
        ["ne"]
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

    numeric_string_relational_coerces_to_number => {
        r#"<?php
echo '07' < 10 ? 'lt' : 'ge';
"#,
        ["lt"]
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
        ["yes"]
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
        ["no"]
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

    coalescing_ternary_interaction => {
        r#"<?php
$a = '';
echo ($a ?: 'fallback') . '|';
$b = null;
echo ($b ?? 'fallback');
"#,
        ["fallback|fallback"]
    };

    additive_vs_ternary_precedence => {
        r#"<?php
echo 1 + 2 ? 3 : 4;
echo '|';
echo (1 + 2) ? 3 : 4;
"#,
        ["3|3"]
    };

    and_word_vs_andor_precedence => {
        r#"<?php
$x = true;
$x = false && false;
echo $x ? 'T' : 'F';
echo '|';
$y = false and false;
echo $y ? 'T' : 'F';
"#,
        ["F|F"]
    };

    loose_equality_falsey_chain => {
        r#"<?php
echo (null == 0) ? '1' : '0';
echo '|';
echo ('' == false) ? '1' : '0';
echo '|';
echo ([] == false) ? '1' : '0';
echo '|';
echo (0 == false) ? '1' : '0';
echo '|';
echo (0 === false) ? '1' : '0';
"#,
        ["1|1|1|1|0"]
    };

    string_numeric_spaces_and_zeros => {
        r#"<?php
echo (" 7" == 7) ? '1' : '0';
echo '|';
echo ("07" == "7") ? '1' : '0';
echo '|';
echo ("07" == 7) ? '1' : '0';
echo '|';
echo ("007" === "7") ? '1' : '0';
echo '|';
echo ("0" == 0) ? '1' : '0';
"#,
        ["1|1|1|0|1"]
    };

    string_floating_prefix_plus_suffix => {
        r#"<?php
echo ("2.3" == 2.3) ? '1' : '0';
echo '|';
echo ("2.3" === 2.3) ? '1' : '0';
echo '|';
echo ("+5" == 5) ? '1' : '0';
echo '|';
echo ("5e2" == 500) ? '1' : '0';
echo '|';
echo ("5e2" === 500) ? '1' : '0';
"#,
        ["1|0|1|1|0"]
    };

    array_loose_strict_edge_cases => {
        r#"<?php
echo ([1, 2] == [1, 2]) ? '1' : '0';
echo '|';
echo ([1, 2] === [1, 2]) ? '1' : '0';
echo '|';
echo (['1', '2'] == [1, 2]) ? '1' : '0';
echo '|';
echo (['1', '2'] === [1, 2]) ? '1' : '0';
echo '|';
echo (['a' => 1] == ['a' => 1]) ? '1' : '0';
"#,
        ["1|1|1|0|1"]
    };

    object_identity_and_equality_basics => {
        r#"<?php
$first = new stdClass();
$second = new stdClass();
$alias = $first;
echo ($first == $second) ? '1' : '0';
echo '|';
echo ($first === $second) ? '1' : '0';
echo '|';
echo ($first == $alias) ? '1' : '0';
echo '|';
echo ($first === $alias) ? '1' : '0';
"#,
        ["0|0|1|1"]
    };

    object_to_string_numeric_comparison => {
        r#"<?php
class Box {
    public function __toString(): string { return '7'; }
}
$box = new Box();
echo ($box == 7) ? '1' : '0';
echo '|';
echo ($box === 7) ? '1' : '0';
echo '|';
echo ($box == '7') ? '1' : '0';
echo '|';
echo ((string)$box == 7) ? '1' : '0';
"#,
        ["1|0|1|1"]
    };

    settype_invalid_to_invalid => {
        r#"<?php
$value = 'abc';
settype($value, 'integer');
echo $value . '|';
$value = 1.9;
settype($value, 'string');
echo $value;
"#,
        ["0|1.9"]
    };

    cast_chain_runtime_edge => {
        r#"<?php
echo (int)'12abc' . '|';
echo (float)'12abc' . '|';
echo (bool)'0' . '|';
echo (bool)'false' . '|';
echo (string)false . '|';
echo (string)true;
"#,
        ["12|12|0|1||1"]
    };

    is_bool_and_bool_as_string => {
        r#"<?php
echo is_bool('true') ? 't1' : 't0';
echo '|';
echo is_bool((bool)'false') ? 'c1' : 'c0';
echo '|';
echo ((bool)0 === false) ? 'e1' : 'e0';
echo '|';
echo ((bool)'0' === false) ? 'f1' : 'f0';
"#,
        ["t0|c1|e1|f1"]
    };
}
