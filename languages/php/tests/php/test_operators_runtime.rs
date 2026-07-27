//! Compound assignment, null coalescing assignment, exponent, and silence operator.

crate::php_cases! {
    concat_assign_appends_string => {
        r#"<?php
$s = 'a';
$s .= 'b';
echo $s;
"#,
        ["ab"]
    };

    plus_assign_accumulates_integer => {
        r#"<?php
$n = 1;
$n += 4;
echo $n;
"#,
        ["5"]
    };

    null_coalesce_assign_sets_when_null => {
        r#"<?php
$a = null;
$a ??= 'default';
echo $a;
"#,
        ["default"]
    };

    null_coalesce_assign_skips_when_set => {
        r#"<?php
$a = 'keep';
$a ??= 'replace';
echo $a;
"#,
        ["keep"]
    };

    exponent_assign_squares_in_place => {
        r#"<?php
$n = 3;
$n **= 2;
echo $n;
"#,
        ["9"]
    };

    bitwise_or_assign_sets_flag => {
        r#"<?php
$m = 0b001;
$m |= 0b010;
echo decbin($m);
"#,
        ["11"]
    };

    bitwise_and_assign_clears_bits => {
        r#"<?php
$m = 0b111;
$m &= 0b101;
echo decbin($m);
"#,
        ["101"]
    };

    xor_assign_toggles_bits => {
        r#"<?php
$m = 0b1010;
$m ^= 0b1100;
echo decbin($m);
"#,
        ["110"]
    };

    shift_left_assign_doubles => {
        r#"<?php
$n = 3;
$n <<= 1;
echo $n;
"#,
        ["6"]
    };

    shift_right_assign_halves => {
        r#"<?php
$n = 8;
$n >>= 1;
echo $n;
"#,
        ["4"]
    };

    silence_operator_suppresses_undefined_variable_notice => {
        r#"<?php
echo @$missing ?? 'fallback';
"#,
        ["fallback"]
    };

    ternary_short_form_selects_truthy_branch => {
        r#"<?php
echo 1 ? 'yes' : 'no';
"#,
        ["yes"]
    };

    elvis_operator_returns_first_non_empty => {
        r#"<?php
echo '' ?: 'fallback';
"#,
        ["fallback"]
    };

    identity_vs_equality_int_and_string => {
        r#"<?php
echo (1 == '1') ? 'eq' : 'ne';
echo (1 === '1') ? 'id' : 'nid';
"#,
        ["eqnid"]
    };

    modulo_assign_wraps_counter => {
        r#"<?php
$n = 7;
$n %= 4;
echo $n;
"#,
        ["3"]
    };

    arithmetic_precedence_mul_before_add => {
        r#"<?php
echo 1 + 2 * 3;
echo 3 * (1 + 2);
"#,
        ["79"]
    };

    null_coalesce_vs_ternary_precedence => {
        r#"<?php
$user = null;
$fallback = 'default';
echo $user ?? 'fallback';
echo $user ?: $fallback;
"#,
        ["fallbackdefault"]
    };

    identity_of_nested_arrays => {
        r#"<?php
$a = ['x' => 1];
$b = ['x' => 1];
$c = $a;
echo ($a == $b) ? 'eq' : 'neq';
echo ($a === $c) ? 'id' : 'nid';
echo ($a === $b) ? 'same' : 'diff';
"#,
        ["eqidsame"]
    };

    spaceship_chain => {
        r#"<?php
echo (1 <=> 1);
echo (2 <=> 1);
echo (1 <=> 2);
"#,
        ["01-1"]
    };

    boolean_precedence_and_short_circuit => {
        r#"<?php
$hits = 0;
if (false && ++$hits) { }
if (true && false || $hits === 1) { echo 'ok'; }
echo $hits;
"#,
        ["0"]
    };

    increment_by_reference_noting => {
        r#"<?php
$arr = [1, 2, 3];
$ref = &$arr[0];
$ref++;
echo $arr[0];
"#,
        ["2"]
    };

    mixed_numeric_string_arithmetic => {
        r#"<?php
echo 2 + "3";
echo "10" - 1;
echo "10" / "2";
echo "5" * "2";
"#,
        ["59510"]
    };

    coalesce_assignment_existing_falsey => {
        r#"<?php
$value = 0;
$value ??= 'default';
echo $value;
"#,
        ["0"]
    };

    ternary_precedence_with_plus => {
        r#"<?php
$a = 1;
$b = 2;
echo $a > 0 ? $a + $b : $b + 5;
"#,
        ["3"]
    };

    null_coalescing_without_parentheses_keeps_key_lookup => {
        r#"<?php
$cfg = ['db' => ['host' => null], 'fallback' => '127.0.0.1'];
echo $cfg['db']['host'] ?? $cfg['fallback'];
"#,
        ["127.0.0.1"]
    };

    spaceship_numeric_and_string_compare => {
        r#"<?php
echo 3 <=> 4;
echo 5 <=> 5;
echo 'b' <=> 'a';
"#,
        ["-101"]
    };

    ternary_nested_right_assoc => {
        r#"<?php
echo true ? (1 ? 'a' : 'b') : 'c';
echo false ? 'x' : (false ? 'y' : 'z');
"#,
        ["az"]
    };

    precedence_with_modulo_and_addition => {
        r#"<?php
echo 10 + 9 % 3;
echo 10 * 9 % 3;
"#,
        ["100"]
    };

    nullsafe_coalesce_and_property_read => {
        r#"<?php
class Node {
    public ?Node $next = null;
}
$head = new Node();
echo $head?->next?->name ?? 'none';
"#,
        ["none"]
    };

    bitwise_not_flips_integer_bits => {
        r#"<?php
echo ~0 & 7;
"#,
        ["7"]
    };

    power_operator_has_lower_than_multiplication => {
        r#"<?php
echo 2 * 2 ** 3;
echo (2 * 2) ** 3;
"#,
        ["1664"]
    };

    increment_prefix_postfix_diff => {
        r#"<?php
$i = 1;
$a = ++$i;
$b = $i++;
echo $a;
echo $b;
echo $i;
"#,
        ["223"]
    };

    assignment_expression_with_reference => {
        r#"<?php
$a = [10, 20];
$first = &$a[0];
$first += 5;
echo $a[0];
"#,
        ["15"]
    };

    boolean_identity_checks => {
        r#"<?php
echo (true === 1) ? 't' : 'f';
echo (false === 0) ? 't' : 'f';
echo (null === null) ? 't' : 'f';
"#,
        ["fft"]
    };

    logical_operator_no_short_circuit_and_operator => {
        r#"<?php
$a = 0;
echo ($a and 1) . ',';
echo ($a && 1) . ',';
echo (1 or 0) . ',';
echo (1 || 0);
"#,
        [",,1,1"]
    };

    string_concatenation_precedence_with_plus_not_supported => {
        r#"<?php
$s = '3' + '4';
$t = '3' . '4';
echo $s;
echo $t;
"#,
        ["734"]
    };

    xor_operator_without_parentheses => {
        r#"<?php
echo 0 xor 0;
echo 0 xor 1;
echo 1 xor 0;
echo 1 xor 1;
"#,
        ["11"]
    };

    and_or_precedence_demo => {
        r#"<?php
echo (0 && 1) . ',';
echo 0 and 1;
echo (0 || 1) . ',';
echo 0 or 1;
"#,
        [",1,1"]
    };

    and_vs_andand_assignment_precedence => {
        r#"<?php
$a = true and false;
echo $a ? 'T' : 'F';
echo '|';
$b = true && false;
echo $b ? 'T' : 'F';
echo '|';
$c = false or true;
echo $c ? 'T' : 'F';
echo '|';
$d = false || true;
echo $d ? 'T' : 'F';
"#,
        ["T|F|F|T"]
    };

    shift_arithmetic_precedence => {
        r#"<?php
echo (1 << 2) + 1;
echo '|';
echo 1 << 2 + 1;
echo '|';
echo (1 + 2) << 1;
echo '|';
echo 1 + (2 << 1);
"#,
        ["5|8|6|5"]
    };

    null_coalesce_falsey_values_runtime => {
        r#"<?php
$name = '';
echo (($name ?? 'fallback') === '' ? 'left' : 'right') . '|';
echo (($name ?: 'fallback') === 'fallback' ? 'left' : 'right');
"#,
        ["left|left"]
    };

    truthiness_and_equality_chain_runtime => {
        r#"<?php
echo (1 == '1') ? 'S' : 'D';
echo (1 === '1') ? 'S' : 'D';
echo (0 == false) ? 'S' : 'D';
echo (0 === false) ? 'S' : 'D';
echo ([] == false) ? 'S' : 'D';
echo ([] === false) ? 'S' : 'D';
"#,
        ["SDSDSD"]
    };

    nullsafe_ternary_and_coalesce_interaction_runtime => {
        r#"<?php
class Holder {
    public function value(): ?string { return null; }
}
$obj = new Holder();
echo ($obj->value() ?? 'missing') . '|';
echo ($obj->value() ?: 'fallback') . '|';
$obj2 = null;
echo ($obj2?->value() ?? 'obj-null');
"#,
        ["missing|fallback|obj-null"]
    };

    precedence_of_plus_with_andand_runtime => {
        r#"<?php
echo 1 + 2 && true ? 'A' : 'B';
echo '|';
echo (1 + 2) && true ? 'A' : 'B';
echo '|';
echo (1 + 2) > 1 && true ? 'A' : 'B';
echo '|';
echo (1 + 2 > 1) && false ? 'A' : 'B';
"#,
        ["A|A|A|B"]
    };

    logical_not_precedence_with_and => {
        r#"<?php
echo !false && true ? 'A' : 'B';
echo '|';
echo !(false && true) ? 'A' : 'B';
echo '|';
echo !('a' === 'b' || 'a' === 'a') ? 'Y' : 'N';
"#,
        ["A|A|N"]
    };

    coalesce_right_associative_chain_runtime => {
        r#"<?php
echo (null ?? null ?? 'x');
echo '|';
echo (0 ?? 'zero');
echo '|';
echo ('' ?? 'fallback');
"#,
        ["x|0|"]
    };

    equality_negation_and_identity_runtime => {
        r#"<?php
echo (1 != '1') ? 'N' : 'S';
echo '|';
echo (1 !== '1') ? 'N' : 'S';
echo '|';
echo (null != '') ? 'N' : 'S';
echo '|';
echo (null !== '');
"#,
        ["S|N|S|1"]
    };

    null_coalesce_and_ternary_grouping_runtime => {
        r#"<?php
$value = null;
echo ($value ?? 'null-fallback');
echo '|';
echo $value ? 'truthy' : ($value ?? 'or-default');
echo '|';
$value = '';
echo $value ?? 'none';
echo '|';
echo $value ?: 'fallback';
"#,
        ["null-fallback|or-default||fallback"]
    };

    ternary_is_right_associative_runtime => {
        r#"<?php
echo false ? 'a' : false ? 'b' : 'c';
echo '|';
echo true ? (false ? 'x' : 'y') : 'z';
"#,
        ["c|y"]
    };

    null_coalesce_right_associative_chain_runtime => {
        r#"<?php
$a = null;
$b = null;
$c = 'active';
echo $a ?? $b ?? $c;
echo '|';
$data = ['x' => null, 'y' => 'v', 'z' => 'ignore'];
echo ($data['x'] ?? $data['y'] ?? $data['z']);
"#,
        ["active|v"]
    };

    null_coalesce_skips_empty_string_runtime => {
        r#"<?php
$value = '';
echo ($value ?? 'fallback-1');
echo '|';
echo ($value ?: 'fallback-2');
echo '|';
$value = 0;
echo ($value ?? 'fallback-3');
echo '|';
echo ($value ?: 'fallback-4');
"#,
        ["|fallback-2|0|fallback-4"]
    };

    spaceship_operator_non_scalar_operands_runtime => {
        r#"<?php
echo ([1, 2] <=> [1, 3]);
echo '|';
echo (['a' => 1] <=> ['a' => 1]);
echo '|';
echo (new stdClass() <=> new stdClass());
"#,
        ["-1|0|0"]
    };

    spaceship_operator_references_same_object_runtime => {
        r#"<?php
$obj = new stdClass();
$alias = $obj;
echo ($obj <=> $alias);
echo '|';
echo ($obj === $alias) ? 'same' : 'not';
"#,
        ["0|same"]
    };

    logical_keyword_and_symbol_precedence_runtime => {
        r#"<?php
$x = false;
echo ($x and true || true) ? 't' : 'f';
echo '|';
echo ($x && true || true) ? 't' : 'f';
echo '|';
$y = false;
$y = false and true;
echo $y ? 't' : 'f';
echo '|';
$z = false && true;
echo $z ? 't' : 'f';
"#,
        ["f|t|f|f"]
    };

    bitwise_shift_then_bitwise_runtime => {
        r#"<?php
echo (1 << 3) & 14;
echo '|';
echo 1 << (3 & 2);
echo '|';
echo (8 >> 1) ^ 3;
echo '|';
echo (9 >> 1) | 2;
"#,
        ["8|4|7|6"]
    };

    combined_numeric_string_equality_runtime => {
        r#"<?php
echo ('01' == 1) ? 'eq' : 'ne';
echo '|';
echo ('01' === 1) ? 'eq' : 'ne';
echo '|';
echo ('10' <=> 2);
echo '|';
echo ('2' <=> 10);
"#,
        ["eq|ne|1|-1"]
    };

    and_or_xor_truth_tables_runtime => {
        r#"<?php
echo (0 and 0) ? 'a' : 'f';
echo '|';
echo (0 and 1) ? 'a' : 'f';
echo '|';
echo (1 or 0) ? 'a' : 'f';
echo '|';
echo (1 xor 1) ? 'a' : 'f';
echo '|';
echo (1 xor 0) ? 'a' : 'f';
"#,
        ["f|f|a|f|a"]
    };

    precedence_between_addition_and_null_coalesce_runtime => {
        r#"<?php
echo 1 + 2 + 3 ?? 'fallback';
echo '|';
echo 0 + (null ?? 7);
echo '|';
echo 4 + (null ?? 1) . '';
echo '|';
echo (null ?? 1) + 4;
"#,
        ["6|7|5|5"]
    };

    assignment_and_ternary_precedence_runtime => {
        r#"<?php
$value = 0;
$result = $value ? 'truthy' : 'falsey';
echo $result;
echo '|';
$value = 1;
$result = $value > 0 ? 'gt0' : 'le0';
echo $result;
echo '|';
$value = 0;
echo $value ? 'first' : $value ? 'second' : 'third';
echo '|';
echo ($value ? 'first' : ($value ? 'second' : 'third'));
"#,
        ["falsey|gt0|third|third"]
    };

    boolean_negation_chain_runtime => {
        r#"<?php
echo !0;
echo '|';
echo !'';
echo '|';
echo !!'';
echo '|';
echo !(!'');
echo '|';
echo !true;
echo '|';
echo !!0;
"#,
        ["1|1|0|0|0|0"]
    };

    unary_and_nullsafe_truthiness_runtime => {
        r#"<?php
class Holder {
    public ?string $value = null;
}
$holder = new Holder();
echo ($holder->value ?? 'none') . '|';
echo (!$holder->value) . '|';
echo ($holder->value ?: 'fallback') . '|';
$holder->value = 'x';
echo ($holder->value ?? 'none') . '|';
echo (!$holder->value) . '|';
echo ($holder->value ?: 'fallback');
"#,
        ["none|1|fallback|x|0|x"]
    };

    nullsafe_in_conditional_assign_runtime => {
        r#"<?php
class Child {
    public function name(): ?string { return null; }
}
class ParentObj {
    public ?Child $child = null;
}
$obj = new ParentObj();
echo ($obj->child?->name() ?? 'missing') . '|';
$obj->child = new Child();
echo ($obj->child?->name() ?? 'missing') . '|';
echo (($obj->child?->name() ?: 'fallback'));
"#,
        ["missing|missing|fallback"]
    };

    mixed_assignment_and_short_circuit_runtime => {
        r#"<?php
$count = 0;
$count = $count ?: 1;
echo $count . '|';
$count = 0;
echo ($count && ($count = 9)) . '|';
echo $count . '|';
$count = 1;
echo ($count && ($count = 9)) . '|';
echo $count;
"#,
        ["1||0|1|9"]
    };

    logical_operator_chain_without_parentheses_runtime => {
        r#"<?php
echo true && false || true;
echo '|';
echo true && (false || true);
echo '|';
echo (true && false) || true;
echo '|';
echo false || false || true;
"#,
        ["1|1|1|1"]
    };

    precedence_plus_vs_ternary_runtime => {
        r#"<?php
echo 1 + 2 ? 10 : 20;
echo '|';
echo 1 + (2 ? 10 : 20);
echo '|';
echo 0 + 5 ? 1 : 2;
echo '|';
echo 0 + (5 ? 1 : 2);
"#,
        ["10|11|1|1"]
    };

    null_coalesce_nested_parentheses_runtime => {
        r#"<?php
echo (null ?? 'x') . '|' . ('' ?? 'y') . '|';
$payload = ['a' => null, 'b' => 'y'];
echo ($payload['a'] ?? $payload['b']) . '|';
echo ((null ?? $payload['a']) ?? $payload['b']);
"#,
        ["x||y|y"]
    };
}
