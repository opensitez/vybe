use super::helpers::{compile_ok, run_prints};

fn assert_php_output(src: &str, expected: &[&str]) {
    assert_eq!(
        run_prints(src),
        expected
            .iter()
            .map(|value| (*value).to_string())
            .collect::<Vec<_>>()
    );
}

// Arithmetic
#[test]
fn add_sub_mul_div() {
    compile_ok("<?php $x = 1 + 2 * 3 - 4 / 2;");
}
#[test]
fn modulo() {
    compile_ok("<?php $x = 10 % 3;");
}
#[test]
fn power() {
    compile_ok("<?php $x = 2 ** 10;");
}
#[test]
fn unary_neg() {
    compile_ok("<?php $x = -$a;");
}
#[test]
fn unary_not() {
    compile_ok("<?php $x = !$a;");
}
#[test]
fn unary_bitnot() {
    compile_ok("<?php $x = ~$a;");
}

// String concat
#[test]
fn concat_dot() {
    compile_ok("<?php $x = 'hello' . ' ' . 'world';");
}
#[test]
fn concat_assign() {
    compile_ok("<?php $x = 'a'; $x .= 'b';");
}

// Comparison
#[test]
fn loose_eq() {
    compile_ok("<?php $x = $a == $b;");
}
#[test]
fn loose_ne() {
    compile_ok("<?php $x = $a != $b;");
}
#[test]
fn strict_eq() {
    compile_ok("<?php $x = $a === $b;");
}
#[test]
fn strict_ne() {
    compile_ok("<?php $x = $a !== $b;");
}
#[test]
fn lt_gt_le_ge() {
    compile_ok("<?php $x = $a < $b; $y = $a > $b; $z = $a <= $b; $w = $a >= $b;");
}
#[test]
fn spaceship() {
    compile_ok("<?php $x = 1 <=> 2;");
}

// Logical
#[test]
fn and_or() {
    compile_ok("<?php $x = $a && $b || $c;");
}
#[test]
fn short_circuit_and() {
    compile_ok("<?php $x = false && expensive();");
}
#[test]
fn short_circuit_or() {
    compile_ok("<?php $x = true || expensive();");
}

// Bitwise
#[test]
fn bitwise_ops() {
    compile_ok("<?php $x = $a & $b | $c ^ $d; $y = $a << 2; $z = $b >> 1;");
}

// Ternary / null coalesce
#[test]
fn ternary() {
    compile_ok("<?php $x = $a ? 'yes' : 'no';");
}
#[test]
fn short_ternary() {
    compile_ok("<?php $x = $a ?: 'default';");
}
#[test]
fn null_coalesce() {
    compile_ok("<?php $x = $a ?? 'default';");
}

// Increment / Decrement
#[test]
fn pre_inc() {
    compile_ok("<?php $x = 0; ++$x;");
}
#[test]
fn post_inc() {
    compile_ok("<?php $x = 0; $x++;");
}
#[test]
fn pre_dec() {
    compile_ok("<?php $x = 0; --$x;");
}
#[test]
fn post_dec() {
    compile_ok("<?php $x = 0; $x--;");
}

// Assignment
#[test]
fn assign() {
    compile_ok("<?php $x = 5;");
}
#[test]
fn add_assign() {
    compile_ok("<?php $x = 0; $x += 5;");
}
#[test]
fn sub_assign() {
    compile_ok("<?php $x = 10; $x -= 3;");
}
#[test]
fn mul_assign() {
    compile_ok("<?php $x = 2; $x *= 4;");
}
#[test]
fn div_assign() {
    compile_ok("<?php $x = 10; $x /= 2;");
}
#[test]
fn mod_assign() {
    compile_ok("<?php $x = 10; $x %= 3;");
}
// **= not yet supported by lexer (no StarStarEq token)
// #[test] fn pow_assign() { compile_ok("<?php $x = 2; $x **= 8;"); }
#[test]
fn array_access_assign() {
    compile_ok("<?php $a = [1,2]; $a[0] = 99;");
}
#[test]
fn assoc_access_assign() {
    compile_ok("<?php $a = []; $a['key'] = 'value';");
}
#[test]
fn property_assign() {
    compile_ok("<?php $obj->name = 'test';");
}

#[test]
fn logical_operators_runtime_results() {
    assert_php_output(
        r#"<?php
$_SERVER = [];
if (!isset($defaultLang) && !empty($_SERVER['HTTP_ACCEPT_LANGUAGE'])) {
	echo 'and-bad';
} else {
	echo 'and-ok';
}

if (!empty($_SERVER['HTTP_ACCEPT_LANGUAGE']) || !isset($defaultLang)) {
	echo 'or-ok';
} else {
	echo 'or-bad';
}

if (true and false) {
	echo 'word-and-bad';
} else {
	echo 'word-and-ok';
}

if (false or true) {
	echo 'word-or-ok';
} else {
	echo 'word-or-bad';
}

if (true xor true) {
	echo 'word-xor-bad';
} else {
	echo 'word-xor-ok';
}
"#,
        &["and-okor-okword-and-okword-or-okword-xor-ok"],
    );
}

#[test]
fn arithmetic_comparison_and_control_operator_runtime_results() {
    assert_php_output(
        r#"<?php
echo 1 + 2;
echo 7 - 4;
echo 6 * 7;
echo 7 / 2;
echo 7 % 3;
echo 2 ** 3;
echo 'a' . 'b';
echo (-5) + 8;
echo (+5);
echo (!false) ? 't' : 'f';
echo (2 < 3) ? 't' : 'f';
echo (3 > 2) ? 't' : 'f';
echo (3 <= 3) ? 't' : 'f';
echo (4 >= 5) ? 't' : 'f';
echo (2 == '2') ? 't' : 'f';
echo (2 === '2') ? 't' : 'f';
echo (2 != 3) ? 't' : 'f';
echo (2 !== '2') ? 't' : 'f';
echo 1 <=> 2;
echo 2 <=> 2;
echo 3 <=> 2;
echo null ?? 'fallback';
echo 'value' ?? 'fallback';
echo false ? 'then' : 'else';
echo 0 ?: 'fallback';
"#,
        &["33423.518ab35ttttftftt-101fallbackvalueelsefallback"],
    );
}

#[test]
fn bitwise_and_shift_operator_runtime_results() {
    assert_php_output(
        r#"<?php
echo 6 & 3;
echo 6 | 3;
echo 6 ^ 3;
echo 1 << 3;
echo 8 >> 2;
echo ~1;
"#,
        &["27582-2"],
    );
}

#[test]
fn compound_assignment_operator_runtime_results() {
    assert_php_output(
        r#"<?php
$x = 5;
$x += 2;
echo $x;
$x -= 4;
echo $x;
$x *= 3;
echo $x;
$x /= 9;
echo $x;
$x %= 2;
echo $x;

$text = 'a';
$text .= 'b';
echo $text;

$bits = 6;
$bits &= 3;
echo $bits;
$bits |= 4;
echo $bits;
$bits ^= 1;
echo $bits;

$shift = 1;
$shift <<= 3;
echo $shift;
$shift >>= 2;
echo $shift;

$fallback = null;
$fallback ??= 'set';
echo $fallback;
$fallback ??= 'again';
echo $fallback;
"#,
        &["73911ab26782setset"],
    );
}

#[test]
fn operator_precedence_runtime_results() {
    assert_php_output(
        r#"<?php
echo (1 + 2) * 3;
echo (4 + 6) / 2;
echo 2 ** 3 ** 2;
echo (2 ** 3) ** 2;
echo 1 + 2 * 3;
echo 1 + 2 << 2;
echo (1 + 2) << 2;
echo 8 >> 1 + 1;
echo (8 >> 1) + 1;
echo 3 + 4 * 2 < 20 && 3 < 4;
"#,
        &["955126471212251"],
    );
}

#[test]
fn operator_null_coalesce_and_coalescing_assignment_runtime() {
    assert_php_output(
        r#"<?php
$value = null;
$fallback = $value ?? 'fallback';
$value ??= 99;
$value ??= 11;
echo $fallback . '|' . $value . '|';

$flags = ['a' => null, 'b' => 'set', 'c' => 'ignored'];
echo $flags['a'] ?? $flags['b'];
echo $flags['a'] ?? $flags['c'];
echo $flags['a'] ?? 'end';
"#,
        &["fallback|99|setignoredend"],
    );
}

#[test]
fn logical_assignment_operators_runtime() {
    assert_php_output(
        r#"<?php
$first = true;
$first &&= false;
echo $first ? 'true' : 'false';

$second = false;
$second ||= true;
echo $second ? 'true2' : 'false2';

$third = null;
$third ??= 'seeded';
$third .= '-done';
echo $third;
"#,
        &["falsetrue2seeded-done"],
    );
}

#[test]
fn identity_and_loose_equality_runtime() {
    assert_php_output(
        r#"<?php
echo (1 == '1') ? 'eq1' : 'neq1';
echo (1 === '1') ? 'eq2' : 'neq2';
echo (0 == false) ? 'eq3' : 'neq3';
echo (0 === false) ? 'eq4' : 'neq4';
echo (null == false) ? 'eq5' : 'neq5';
echo (null === false) ? 'eq6' : 'neq6';
"#,
        &["eq1neq2eq3neq4eq5neq6"],
    );
}

#[test]
fn array_union_preserves_left_hand_side_runtime() {
    assert_php_output(
        r#"<?php
$left = ['first' => 'l', 2 => 'two'];
$right = [1 => 'one', 2 => 'override', 'extra' => 'x'];
$merged = $left + $right;
echo $merged['first'];
echo $merged[2];
echo $merged[1];
echo array_key_exists('extra', $merged) ? 'extra' : 'noextra';
"#,
        &["ltwooneextra"],
    );
}

#[test]
fn null_coalesce_and_ternary_precedence_runtime() {
    assert_php_output(
        r#"<?php
$value = null;
echo (($value ?? 'a') ?: 'fallback') . '|';
$value = '';
echo (($value ?? 'a') ?: 'fallback') . '|';
$value = 0;
echo (($value ?? 'a') ?: 'fallback') . '|';
echo (($value ?? null) ?: 'fallback') . '|';
"#,
        &["a|fallback|fallback|fallback|"],
    );
}

#[test]
fn concat_then_addition_precedence_runtime() {
    assert_php_output(
        r#"<?php
echo 'a' . 'b' . 'c';
echo ('3' . 2) + 1;
echo '3' . (2 + 1);
echo 'x' . (2 + 3);
"#,
        &["abc3333x5"],
    );
}

#[test]
fn arithmetic_with_float_and_int_mix_runtime() {
    assert_php_output(
        r#"<?php
echo 7 / 2;
echo '|';
echo 7.0 + 1;
echo '|';
echo 8 / 2.0;
"#,
        &["3.5|8|4"],
    );
}

#[test]
fn comparison_chain_on_strings_and_numbers_runtime() {
    assert_php_output(
        r#"<?php
echo ('2' < '10') ? 'strlt' : 'strgt';
echo '|';
echo (2 < '10') ? 'nlt' : 'ngt';
echo '|';
echo (2 <=> '2');
"#,
        &["strlt|nlt|0"],
    );
}

#[test]
fn logical_xor_and_word_operator_interaction_runtime() {
    assert_php_output(
        r#"<?php
echo (true xor false) ? 'x1' : 'x0';
echo '|';
echo (true and false) ? 'and1' : 'and0';
echo '|';
echo (true or false) ? 'or1' : 'or0';
"#,
        &["x1|and0|or1"],
    );
}

#[test]
fn nullsafe_operator_with_coalesce_in_chain_runtime() {
    assert_php_output(
        r#"<?php
class Child {
    public string $value = 'ok';
    public function getChild(): ?Child { return null; }
}
class ParentObj {
    public Child $child;
    public function __construct() { $this->child = new Child(); }
}
$obj = new ParentObj();
echo $obj->child?->value . '|';
echo $obj->child->getChild()?->value ?? 'none';
"#,
        &["ok|none"],
    );
}

#[test]
fn bitwise_alias_and_parentheses_runtime() {
    assert_php_output(
        r#"<?php
echo (3 | 1) & 2;
echo '|';
echo (7 ^ 2) >> 1;
echo '|';
echo (12 << 1) | 1;
"#,
        &["2|2|25"],
    );
}

#[test]
fn nullsafe_with_parentheses_and_default_runtime() {
    assert_php_output(
        r#"<?php
class Holder {
    public function value(): ?self { return null; }
}
class Node {
    public ?Node $next;
    public Holder $holder;
    public function __construct() {
        $this->next = null;
        $this->holder = new Holder();
    }
}
$node = new Node();
$first = $node->next?->holder?->value();
echo $first ?? 'empty';
$second = $node->holder?->value() ?? 'none';
echo '|';
echo $second;
"#,
        &["empty|none"],
    );
}

#[test]
fn error_control_operator_with_undefined_variable_runtime() {
    assert_php_output(
        r#"<?php
echo @$undefined_variable;
echo '|';
echo isset(@$undefined_variable) ? 'defined' : 'not';
"#,
        &["|not"],
    );
}

#[test]
fn spaceship_with_boolean_and_int_runtime() {
    assert_php_output(
        r#"<?php
echo (true <=> false);
echo '|';
echo (false <=> true);
echo '|';
echo (true <=> true);
echo '|';
echo (10 <=> -10);
"#,
        &["1|-1|0|1"],
    );
}

#[test]
fn precedence_unary_minus_and_exponent_runtime() {
    assert_php_output(
        r#"<?php
echo -2 ** 3;
echo '|';
echo (-2) ** 3;
echo '|';
echo 2 ** 3 ** 2;
"#,
        &["-8|-8|512"],
    );
}

#[test]
fn coercive_comparison_string_number_runtime() {
    assert_php_output(
        r#"<?php
echo ("0" == 0) ? 'z' : 'nz';
echo '|';
echo ("0" === 0) ? 'zs' : 'nzs';
echo '|';
echo ("10" < 2) ? 'l' : 'g';
echo '|';
echo ("2" > 10) ? 'g2' : 'l2';
"#,
        &["z|nzs|g|l2"],
    );
}

#[test]
fn bitwise_assignment_and_rebind_runtime() {
    assert_php_output(
        r#"<?php
$bits = 0b1111;
$bits &= 0b1010;
echo $bits;
echo '|';
$bits |= 0b0101;
echo $bits;
echo '|';
$bits ^= 0b0011;
echo $bits;
$bits <<= 1;
echo '|';
echo $bits;
$bits >>= 2;
echo '|';
echo $bits;
"#,
        &["10|15|12|24|6"],
    );
}

#[test]
fn assignment_with_reference_like_chaining_runtime() {
    assert_php_output(
        r#"<?php
$values = ['a' => ['b' => ['c' => 1]]];
$values['a']['b']['c'] += 4;
echo $values['a']['b']['c'];
echo '|';
$copy = $values;
$copy['a']['b']['c'] *= 2;
echo $copy['a']['b']['c'];
echo '|';
echo $values['a']['b']['c'];
"#,
        &["5|10|5"],
    );
}

#[test]
fn subtraction_assignment_with_negative_literals_runtime() {
    assert_php_output(
        r#"<?php
$value = 10;
$value -= -5;
echo $value;
echo '|';
$value += -2;
echo $value;
echo '|';
echo -$value + 3;
"#,
        &["15|13|-10"],
    );
}

#[test]
fn strict_type_juggling_and_string_comparison_runtime() {
    assert_php_output(
        r#"<?php
echo (0 == '0') ? 'eq' : 'ne';
echo '|';
echo (0 === '0') ? 'seq' : 'sne';
echo '|';
echo ('10' > '2') ? 'strgt' : 'strlt';
echo '|';
echo (10 <=> '2');
"#,
        &["eq|sne|strgt|1"],
    );
}

#[test]
fn null_coalesce_falsey_values_runtime() {
    assert_php_output(
        r#"<?php
echo ('' ?? 'empty') . '|';
echo (0 ?? 'zero') . '|';
$value = null;
echo (($value ?? null) ?? 'fallback') . '|';
$value = 0;
echo ($value ?? 'fallback') . '|';
$value = '';
echo ($value ?? 'fallback') . '|';
"#,
        &["|0|fallback|0||"],
    );
}

#[test]
fn logical_xor_with_string_operands_runtime() {
    assert_php_output(
        r#"<?php
echo ((bool)'x' xor (bool) '') ? 'xor1' : 'xor0';
echo '|';
echo ((bool)'x' xor (bool)'y') ? 'xor1' : 'xor0';
"#,
        &["xor1|xor0"],
    );
}

#[test]
fn modulo_with_negative_operands_runtime() {
    assert_php_output(
        r#"<?php
echo 7 % 3;
echo '|';
echo -7 % 3;
echo '|';
echo 7 % -3;
echo '|';
echo -7 % -3;
"#,
        &["1|-1|1|-1"],
    );
}

#[test]
fn bitwise_string_operands_runtime() {
    assert_php_output(
        r#"<?php
echo ord(('a' | 'b')[0]) . '|';
echo ord(('foo' & 'bar')[0]) . ord(('foo' & 'bar')[1]) . ord(('foo' & 'bar')[2]) . '|';
echo ord(('abc' ^ 'bcd')[0]) . ord(('abc' ^ 'bcd')[1]) . ord(('abc' ^ 'bcd')[2]);
"#,
        &["99|989798|317"],
    );
}

#[test]
fn shift_precedence_with_arithmetic_runtime() {
    assert_php_output(
        r#"<?php
echo (3 + 4) << 2;
echo '|';
echo 1 + (2 << 2);
echo '|';
echo (8 >> 1) + (2 << 1);
echo '|';
echo (8 >> 1) + 1 + 0;
"#,
        &["28|9|8|5"],
    );
}

#[test]
fn comparison_with_arrays_and_objects_runtime() {
    assert_php_output(
        r#"<?php
echo ([] == []) ? 'array-eq' : 'array-ne';
echo '|';
echo ([] === []) ? 'array-ident' : 'array-not-ident';
echo '|';
class O {}
$a = new O();
$b = new O();
echo ($a == $b) ? 'obj-eq' : 'obj-ne';
echo '|';
echo ($a === $b) ? 'obj-id' : 'obj-not-id';
"#,
        &["array-eq|array-ident|obj-eq|obj-not-id"],
    );
}

#[test]
fn instanceof_basic_runtime_operator() {
    assert_php_output(
        r#"<?php
class Base {}
class Child extends Base {}
$base = new Base();
$child = new Child();
echo $base instanceof Base ? 'b1' : 'b0';
echo '|';
echo $child instanceof Base ? 'c1' : 'c0';
echo '|';
echo $base instanceof Child ? 'd1' : 'd0';
"#,
        &["b1|c1|d0"],
    );
}

#[test]
fn nullsafe_chain_with_coalesce_and_truthy_checks_runtime() {
    assert_php_output(
        r#"<?php
class Leaf {
    public function value(): ?string {
        return null;
    }
}
class Branch {
    public ?Leaf $leaf;
    public function __construct(?Leaf $leaf) { $this->leaf = $leaf; }
    public function child(): ?Leaf { return $this->leaf; }
}
$withLeaf = new Branch(new Leaf());
$withoutLeaf = new Branch(null);
echo ($withLeaf->child()?->value() ?? 'none1') . '|';
echo ($withoutLeaf->child()?->value() ?? 'none2');
"#,
        &["none1|none2"],
    );
}

#[test]
fn precedence_pow_vs_add_sub_runtime_operator() {
    assert_php_output(
        r#"<?php
echo 2 + 3 ** 2;
echo '|';
echo (2 + 3) ** 2;
echo '|';
echo 2 ** 3 ** 2;
echo '|';
echo (2 ** 3) ** 2;
"#,
        &["11|25|512|64"],
    );
}

#[test]
fn null_coalesce_assign_operator_runtime() {
    assert_php_output(
        r#"<?php
$value = null;
echo ($value ??= 'fallback') . '|';
$value = 'set';
echo ($value ??= 'fallback') . '|';
echo $value;
"#,
        &["fallback|set|set"],
    );
}

#[test]
fn combined_boolean_bitwise_operators_runtime() {
    assert_php_output(
        r#"<?php
$a = true;
$b = false;
echo ($a and $b) ? 'ab1' : 'ab0';
echo '|';
echo ($a && $b) ? 'aa1' : 'aa0';
echo '|';
echo ($a xor $b) ? 'ax1' : 'ax0';
echo '|';
echo ((1 & 3) === 1) ? 'b1' : 'b0';
"#,
        &["ab0|aa0|ax1|b1"],
    );
}

#[test]
fn strictness_of_equality_chain_runtime() {
    assert_php_output(
        r#"<?php
echo (1 == true) . '|';
echo (1 === true) . '|';
echo ("1" == 1) . '|';
echo ("1" === 1) . '|';
echo (0 == false) . '|';
echo (0 === false);
"#,
        &["1|0|1|0|1|0"],
    );
}

#[test]
fn arithmetic_with_assignment_and_precedence_runtime() {
    assert_php_output(
        r#"<?php
$n = 10;
$n += 2 * 3;
echo $n . '|';
$n -= 4 + 1;
echo $n . '|';
$n *= 2 + 1;
echo $n . '|';
$n /= 2 + 1;
echo $n;
"#,
        &["16|11|33|11"],
    );
}

#[test]
fn string_numeric_coercion_addition_vs_concat_runtime() {
    assert_php_output(
        r#"<?php
echo '1' + 2;
echo '|';
echo '1' . 2;
echo '|';
echo 1 . '2';
"#,
        &["3|12|12"],
    );
}

#[test]
fn null_coalesce_nested_array_indexing_chain_runtime() {
    assert_php_output(
        r#"<?php
$payload = ['user' => ['profile' => null]];
echo $payload['user']['name'] ?? 'anon';
echo '|';
echo ($payload['user']['profile'] ?? $payload['user']['fallback'] ?? 'missing');
echo '|';
echo ($payload['team']['owner'] ?? 'team-unknown');
"#,
        &["anon|missing|team-unknown"],
    );
}

#[test]
fn ternary_with_parentheses_and_nullish_falsey_runtime() {
    assert_php_output(
        r#"<?php
echo (0 ? 'truthy' : 'falsey') . '|';
echo (0 ?: 'fallback') . '|';
echo ((0 == false) ? 'eq' : 'neq') . '|';
echo ((0 === false) ? 'strict-eq' : 'strict-neq');
"#,
        &["falsey|fallback|eq|strict-neq"],
    );
}

#[test]
fn bitwise_with_nested_parentheses_runtime() {
    assert_php_output(
        r#"<?php
echo (3 & 1);
echo (3 | 4);
echo '|';
echo (3 | (1 & 2)) . '|';
echo ((5 ^ 1) & 6);
echo '|';
echo ((1 << 3) >> 1);
"#,
        &["17|3|4|4"],
    );
}

#[test]
fn logical_and_or_xor_keyword_vs_symbol_runtime() {
    assert_php_output(
        r#"<?php
$left = (true and false) ? 't' : 'f';
$right = (true && false) ? 't' : 'f';
echo $left . '|';
echo $right . '|';
echo (true xor false) ? 'x1' : 'x0';
echo '|';
echo ((bool) (0 && 1)) ? 'z1' : 'z0';
"#,
        &["f|f|x1|z0"],
    );
}

#[test]
fn comparison_chain_with_string_numeric_conversion_runtime() {
    assert_php_output(
        r#"<?php
echo ('10' < 2) ? 'lt' : 'ge';
echo '|';
echo ('10' > 2) ? 'gt' : 'le';
echo '|';
echo ('2' < 10) ? 'lt2' : 'ge2';
echo '|';
echo (2 <=> '10');
"#,
        &["ge|gt|lt2|-1"],
    );
}

#[test]
fn spaceship_with_floats_and_integers_runtime() {
    assert_php_output(
        r#"<?php
echo (1.5 <=> 1.25);
echo '|';
echo (1.25 <=> 1.25);
echo '|';
echo (1.25 <=> 2.5);
"#,
        &["1|0|-1"],
    );
}

#[test]
fn modulo_and_division_negative_runtime_cases() {
    assert_php_output(
        r#"<?php
echo 10 % 3;
echo '|';
echo -10 % 3;
echo '|';
echo 10 / -2;
echo '|';
echo -10 / 2;
"#,
        &["1|-1|-5|-5"],
    );
}

#[test]
fn increment_decrement_runtime_precedence() {
    assert_php_output(
        r#"<?php
$n = 1;
echo ++$n;
echo '|';
echo $n++;
echo '|';
echo $n;
echo '|';
echo --$n;
echo '|';
echo $n--;
echo '|';
echo $n;
"#,
        &["2|2|3|2|2|1"],
    );
}

#[test]
fn instanceof_chain_runtime() {
    assert_php_output(
        r#"<?php
class Base {}
class Child extends Base {}
interface Marker {}
class Worker extends Child implements Marker {}
$value = new Worker();
echo ($value instanceof Child) ? 'child' : 'no-child';
echo '|';
echo ($value instanceof Base) ? 'base' : 'no-base';
echo '|';
echo ($value instanceof Marker) ? 'marker' : 'no-marker';
"#,
        &["child|base|marker"],
    );
}

#[test]
fn match_expression_operator_runtime() {
    assert_php_output(
        r#"<?php
$score = 72;
echo match ($score) {
    100 => 'A',
    90 => 'B',
    80, 70 => 'C',
    default => 'F',
};
echo '|';
$code = 'x';
echo match ($code) {
    'x' => 'ex',
    'y' => 'why',
    default => 'other',
};
"#,
        &["F|ex"],
    );
}

#[test]
fn match_with_guarded_conditions_runtime() {
    assert_php_output(
        r#"<?php
$x = 6;
echo match (true) {
    $x < 0 => 'neg',
    $x < 5 => 'small',
    $x >= 5 && $x <= 10 => 'mid',
    default => 'big',
};
"#,
        &["mid"],
    );
}

#[test]
fn match_returning_array_runtime() {
    assert_php_output(
        r#"<?php
$mode = 'json';
$value = match ($mode) {
    'json' => ['kind' => 'json'],
    'text' => 'plain',
    default => null,
};
echo is_array($value) ? $value['kind'] : 'none';
"#,
        &["json"],
    );
}

#[test]
fn spread_operator_for_function_arguments_runtime() {
    assert_php_output(
        r#"<?php
function sum(int ...$nums): int {
    $total = 0;
    foreach ($nums as $n) { $total += $n; }
    return $total;
}
$numbers = [1, 2, 3];
echo sum(...$numbers);
"#,
        &["6"],
    );
}

#[test]
fn spread_operator_array_merge_runtime() {
    assert_php_output(
        r#"<?php
$head = [1, 2];
$tail = [...$head, 3];
echo implode(',', $tail);
"#,
        &["1,2,3"],
    );
}

#[test]
fn coalesce_with_missing_key_in_nested_expression_runtime() {
    assert_php_output(
        r#"<?php
$data = ['a' => ['b' => null]];
echo $data['a']['b'] ?? 'fallback-a';
echo '|';
echo $data['a']['c'] ?? 'fallback-c';
echo '|';
echo ($data['x']['y'] ?? null) ?? 'fallback-xy';
"#,
        &["fallback-a|fallback-c|fallback-xy"],
    );
}

#[test]
fn nullsafe_operator_chain_with_fallback_runtime() {
    assert_php_output(
        r#"<?php
class Child {
    public function value(): ?string {
        return null;
    }
}
class ParentNode {
    public function child(): ?Child {
        return null;
    }
}
$node = new ParentNode();
echo $node->child()?->value() ?? 'none';
"#,
        &["none"],
    );
}

#[test]
fn cast_and_identity_operators_runtime() {
    assert_php_output(
        r#"<?php
echo (int) '12';
echo '|';
echo (int) '12.9';
echo '|';
echo (bool) '';
echo '|';
echo (bool) 'php';
echo '|';
echo (string) 12;
"#,
        &["12|12||1|12"],
    );
}

#[test]
fn truthiness_matrix_runtime() {
    assert_php_output(
        r#"<?php
echo (((bool) null) === true) ? 'T' : 'F';
echo '|';
echo (((bool) false) === true) ? 'T' : 'F';
echo '|';
echo (((bool) true) === true) ? 'T' : 'F';
echo '|';
echo (((bool) 0) === true) ? 'T' : 'F';
echo '|';
echo (((bool) 1) === true) ? 'T' : 'F';
echo '|';
echo (((bool) -3) === true) ? 'T' : 'F';
echo '|';
echo (((bool) 0.0) === true) ? 'T' : 'F';
echo '|';
echo (((bool) 1.2) === true) ? 'T' : 'F';
echo '|';
echo (((bool) '') === true) ? 'T' : 'F';
echo '|';
echo (((bool) '0') === true) ? 'T' : 'F';
echo '|';
echo (((bool) '1') === true) ? 'T' : 'F';
echo '|';
echo (((bool) 'PHP') === true) ? 'T' : 'F';
echo '|';
echo (((bool) []) === true) ? 'T' : 'F';
echo '|';
echo (((bool) [0]) === true) ? 'T' : 'F';
echo '|';
echo (((bool) ['']) === true) ? 'T' : 'F';
echo '|';
echo (empty([1, 2]) ? 'T' : 'F');
"#,
        &["F|F|T|F|T|T|F|T|F|F|T|T|F|T|T|F"],
    );
}

#[test]
fn equality_variants_runtime() {
    assert_php_output(
        r#"<?php
echo (1 == '1') ? '1' : '0';
echo (1 === '1') ? '1' : '0';
echo (0 == false) ? '1' : '0';
echo (0 === false) ? '1' : '0';
echo ('0' == false) ? '1' : '0';
echo ('0' === false) ? '1' : '0';
echo ('' == false) ? '1' : '0';
echo ('' === false) ? '1' : '0';
echo ([] == false) ? '1' : '0';
echo ([] === false) ? '1' : '0';
echo ([] == null) ? '1' : '0';
echo ([1,2] == [1,2]) ? '1' : '0';
echo ([1,2] === [1,2]) ? '1' : '0';
echo ([1,2] === [2,1]) ? '1' : '0';
echo ('a' != 'b') ? '1' : '0';
echo ('a' <> 'a') ? '1' : '0';
echo (new stdClass() == new stdClass()) ? '1' : '0';
echo (new stdClass() === new stdClass()) ? '1' : '0';
"#,
        &["101010101011101010"],
    );
}

#[test]
fn not_equal_operator_alias_and_negation_runtime() {
    assert_php_output(
        r#"<?php
echo (1 <> 2) ? 'T' : 'F';
echo '|';
echo (1 != 1) ? 'T' : 'F';
echo '|';
echo !true ? 'T' : 'F';
echo '|';
echo !false ? 'T' : 'F';
echo '|';
echo !!false ? 'T' : 'F';
"#,
        &["T|F|F|T|F"],
    );
}

#[test]
fn operator_truthy_checks_inside_conditions_runtime() {
    assert_php_output(
        r#"<?php
function truthy_label(mixed $value): string {
    return $value ? 'T' : 'F';
}
$inputs = [null, 0, 1, '', '0', 'ok', [], [1], false, true];
$out = '';
	foreach ($inputs as $value) {
	    $out .= truthy_label($value);
	}
    echo $out;
"#,
        &["FFTFFTFTFT"],
    );
}

#[test]
fn operator_precedence_full_coverage_runtime() {
    assert_php_output(
        r#"<?php
echo (1 + 2 * 3 - 4 / 2) . '|';
echo ((1 + 2) * 3) . '|';
echo (1 + 2 * 3) . '|';
echo (-2 ** 3) . '|';
echo ((-2) ** 3) . '|';
echo (2 ** 3 ** 2) . '|';
echo ((2 ** 3) ** 2) . '|';
echo (1 + 2 << 1) . '|';
echo (1 + (2 << 1)) . '|';
echo (3 + 4 << 2) . '|';
echo (3 + (4 << 2)) . '|';
echo (7 & 3 | 1) . '|';
echo (7 ^ 3 & 1) . '|';
echo (7 | 3 ^ 1) . '|';
echo ('a' . 1 + 2) . '|';
echo ('a' . (1 + 2)) . '|';
echo (1 < 2 && 2 < 3 ? 'T' : 'F') . '|';
echo (1 < 2 || 2 < 1 ? 'T' : 'F') . '|';
echo (false and true || true ? 'T' : 'F') . '|';
echo (true && false || true ? 'T' : 'F') . '|';
echo (true and false || true ? 'T' : 'F') . '|';
$a = true;
$a = true && false;
echo (($a === false) ? 'F0' : 'T0') . '|';
$a = true;
$a = true and false;
echo (($a === true) ? 'T1' : 'F1') . '|';
echo (0 ?: 2 + 3) . '|';
echo (1 ?: 2 + 3) . '|';
echo (1 ? 2 : 3 + 4) . '|';
echo (0 ? 2 : 3 + 4) . '|';
$payload = ['user' => ['name' => null], 'fallback' => ['name' => 'x']];
echo ($payload['user']['name'] ?? $payload['fallback']['name'] ?? 'none') . '|';
echo (($payload['user']['name'] ?? $payload['fallback']['name']) ?? 'none') . '|';
echo ((0 == false) ? 'T' : 'F') . '|';
echo ((0 === false) ? 'T' : 'F') . '|';
echo ((1 == '1') ? 'T' : 'F') . '|';
echo ((1 === '1') ? 'T' : 'F') . '|';
echo ((1 <=> '1') <=> 0) . '|';
echo (false or true xor true && false ? 'T' : 'F');
"#,
        &["5|9|7|-8|-8|512|64|6|5|28|19|3|6|7|a3|a3|T|T||T|1|F0|T1|5|1|2|7|x|x|T|F|T|F|0|1"],
    );
}

#[test]
fn null_coalescing_and_ternary_edge_precedence_runtime() {
    assert_php_output(
        r#"<?php
echo (null ?? 'fallback') . '|';
echo (0 ?? 99) . '|';
$first = null;
$second = 0;
echo (($first ?? 12) ?: ($second ?? 34));
echo '|';
echo (($second ?? 56) ?: 78);
"#,
        &["fallback|0|12|78"],
    );
}

#[test]
fn equality_juggling_edge_matrix_runtime() {
    assert_php_output(
        r#"<?php
echo ('00' == 0) ? '1' : '0';
echo '|';
echo ('00' === 0) ? '1' : '0';
echo '|';
echo ('01' == 1) ? '1' : '0';
echo '|';
echo (' 1' == 1) ? '1' : '0';
echo '|';
echo ('1e3' == 1000) ? '1' : '0';
echo '|';
echo ('1e3' === 1000) ? '1' : '0';
echo '|';
echo (["1"] == ["1"]) ? '1' : '0';
echo '|';
echo ([1, 2] == [1, '2']) ? '1' : '0';
echo '|';
echo ([1, 2] === [1, '2']) ? '1' : '0';
"#,
        &["1|0|1|1|1|0|1|1|0"],
    );
}

#[test]
fn comparison_operators_short_circuit_side_effects_runtime() {
    assert_php_output(
        r#"<?php
$log = [];
$right = function() use (&$log): bool {
    $log[] = 'right';
    return false;
};

if (false && $right()) {
    echo 'bad';
}
echo (count($log) === 0 ? 'no-right' : 'right-called');
echo '|';
$log = [];
if (true || $right()) {
    echo 'skip';
}
echo (count($log) === 0 ? 'no-right' : 'right-called');
echo '|';
echo ($right() && true) ? 'bad' : 'ok';
echo '|';
echo count($log);
"#,
        &["no-right|skipno-right|ok|1"],
    );
}

#[test]
fn match_exhaustive_default_runtime() {
    assert_php_output(
        r#"<?php
$value = 3;
echo match ($value) {
    1 => 'one',
    2, 3 => 'two-or-three',
    default => 'other',
};
echo '|';
echo match ($value > 1) {
    false => 'small',
    true => 'big',
};
echo '|';
$list = [1, 2, 3];
echo match (true) {
    in_array($value, $list) => 'present',
    default => 'absent',
};
"#,
        &["two-or-three|big|present"],
    );
}

#[test]
fn identity_vs_equality_on_objects_runtime() {
    assert_php_output(
        r#"<?php
$first = new stdClass();
$second = $first;
$third = new stdClass();
echo ($first == $second) ? 'eq1' : 'ne1';
echo '|';
echo ($first === $second) ? 'id1' : 'ni1';
echo '|';
echo ($first == $third) ? 'eq2' : 'ne2';
echo '|';
echo ($first === $third) ? 'id2' : 'ni2';
"#,
        &["eq1|id1|eq2|ni2"],
    );
}

#[test]
fn precedence_or_precedence_and_assignment_runtime() {
    assert_php_output(
        r#"<?php
$x = false;
$y = 1 + 2 * 3;
echo ($x || $y) ? 'ok' : 'bad';

$a = 0;
$b = $a ||= true;
echo '|';
echo $b ? 'A' : 'B';

$c = null;
$c = false || true;
echo '|';
echo $c ? 'T' : 'F';
"#,
        &["ok|A|T"],
    );
}

#[test]
fn equality_truthiness_matrix_runtime() {
    assert_php_output(
        r#"<?php
echo (0 == false) ? 't' : 'f';
echo (0 === false) ? 't' : 'f';
echo ("" == false) ? 't' : 'f';
echo ("" === false) ? 't' : 'f';
echo ([] == false) ? 't' : 'f';
echo ([] === false) ? 't' : 'f';
echo (null == false) ? 't' : 'f';
echo (null === false) ? 't' : 'f';
echo ("0" == false) ? 't' : 'f';
echo ("0" === false) ? 't' : 'f';
"#,
        &["tftftftftf"],
    );
}

#[test]
fn ternary_nullsafe_and_coalesce_interaction_runtime() {
    assert_php_output(
        r#"<?php
$cfg = null;
$a = $cfg ?: 'fallback';
echo $a . '|';

$cfg = 0;
$a = $cfg ?: 'fallback';
echo $a . '|';

$cfg = null;
$a = ($cfg ?? 'fallback') ?: 'none';
echo $a . '|';

$cfg = 0;
$a = ($cfg ?? 'fallback') ?: 'none';
echo $a . '|';
"#,
        &["fallback|fallback|fallback|none|"],
    );
}

#[test]
fn spaceship_and_parentheses_edge_runtime() {
    assert_php_output(
        r#"<?php
echo (1 <=> 2) . '|';
echo (2 <=> 1) . '|';
echo (2 <=> 2) . '|';
echo (3 + 5 <=> 4 + 1) . '|';
echo (false <=> true) . '|';
echo ((5 < 3) <=> (2 < 4));
"#,
        &["-1|1|0|1|-1|-1"],
    );
}

#[test]
fn mixed_operator_chain_with_parentheses_runtime() {
    assert_php_output(
        r#"<?php
echo (1 + 2) * (3 + 4);
echo '|';
echo 1 + (2 * (3 + 4));
echo '|';
echo (2 + 3) ** (2 + 1);
echo '|';
echo ((2 + 3) ** 2) * 2;
"#,
        &["21|15|125|50"],
    );
}

#[test]
fn unary_minus_and_pre_increment_precedence_runtime() {
    assert_php_output(
        r#"<?php
$x = 3;
$y = -$x;
echo $y;
echo '|';
$y = -2;
echo ++$y;
"#,
        &["-3|-1"],
    );
}

#[test]
fn logical_operator_short_circuit_side_effect_mutation_runtime() {
    assert_php_output(
        r#"<?php
$hits = [];
$left = function() use (&$hits) {
    $hits[] = 'left';
    return false;
};
$right = function() use (&$hits) {
    $hits[] = 'right';
    return true;
};
echo false && $left();
echo '|';
echo 0 || $right();
echo '|';
echo count($hits);
echo '|';
echo implode(',', $hits);
"#,
        &["|1|1|right"],
    );
}

#[test]
fn ternary_with_computed_condition_runtime() {
    assert_php_output(
        r#"<?php
echo ((1 + 1) === 2 ? 'eq' : 'ne') . '|';
echo ((1 + 1) === 3 ? 'eq' : 'ne') . '|';
echo ((true ? 1 : 0) ? 't' : 'f') . '|';
echo ((false ? 1 : 2) ? 't' : 'f');
"#,
        &["eq|ne|t|t"],
    );
}

#[test]
fn coalesce_precedence_with_arrays_and_nested_runtime() {
    assert_php_output(
        r#"<?php
$cfg = ['a' => ['b' => null], 'fallback' => 'ok'];
echo ($cfg['a']['b'] ?? $cfg['fallback']);
echo '|';
echo (($cfg['a']['b'] ?? null) ?? $cfg['fallback']);
echo '|';
echo (null ?? $cfg['fallback']);
echo '|';
echo (0 ?? $cfg['fallback']);
"#,
        &["ok|ok|ok|0"],
    );
}

#[test]
fn unary_plus_and_unary_minus_runtime_values() {
    assert_php_output(
        r#"<?php
echo +5 . '|';
echo +(-7) . '|';
echo +('8') . '|';
echo -(-3) . '|';
echo -(1 + 2) . '|';
echo +3.5;
"#,
        &["5|-7|8|3|-3|3.5"],
    );
}

#[test]
fn post_increment_in_expression_chain_runtime() {
    assert_php_output(
        r#"<?php
$counter = 1;
$left = $counter++;
$right = ++$counter;
echo $left . '|' . $right . '|' . $counter . '|' . ($left + $right);
"#,
        &["1|3|3|4"],
    );
}

#[test]
fn bitwise_and_shift_precedence_runtime() {
    assert_php_output(
        r#"<?php
echo (3 & 6 | 9) . '|';
echo ((3 & 6) | 9) . '|';
echo (3 & (6 | 9)) . '|';
echo (1 << 2 | 3) . '|';
echo (1 | 2 << 3);
"#,
        &["11|11|3|7|17"],
    );
}

#[test]
fn word_operators_and_symbol_operators_runtime() {
    assert_php_output(
        r#"<?php
$value = true;
if ($value and false or true) {
    echo 'word-and-or-1';
}
echo '|';
if (($value and false) or true) {
    echo 'word-and-or-2';
}
echo '|';
echo (true or false and false) . '|';
echo ((true or false) and false) . '|';
echo ($value && false or true) . '|';
echo ($value and false || true);
"#,
        &["word-and-or-1|word-and-or-2|1||1|1"],
    );
}

#[test]
fn null_coalescing_and_ternary_edge_interaction_runtime() {
    assert_php_output(
        r#"<?php
$value = null;
echo ($value ?? 'fallback') . '|';
$value = 0;
echo ($value ?? 'fallback') . '|';
echo ($value ?: 'fallback') . '|';
$value = '';
echo ($value ?? 'fallback') . '|';
echo ($value ?: 'fallback') . '|';
echo ((null ?? 'fallback') ?: 'end') . '|';
$value = false;
echo (($value ?? true) ?: true);
"#,
        &["fallback|0|fallback||fallback|fallback|1"],
    );
}

#[test]
fn null_coalescing_assignment_nested_target_runtime() {
    assert_php_output(
        r#"<?php
$cfg = ['x' => null];
$first = $cfg['x'] ??= 'seeded';
echo $cfg['x'] . '|' . $first . '|';
$cfg['y'] ??= 'auto';
echo $cfg['y'] . '|' . ($cfg['y'] ??= 'ignored');
"#,
        &["seeded|seeded|auto|auto"],
    );
}

#[test]
fn dynamic_instanceof_class_name_runtime() {
    assert_php_output(
        r#"<?php
class ParentClass {}
class ChildClass extends ParentClass {}
$obj = new ChildClass();
$type = ChildClass::class;
echo ($obj instanceof $type) . '|';
$type = ParentClass::class;
echo ($obj instanceof $type) . '|';
$type = stdClass::class;
echo ($obj instanceof $type ? 'yes' : 'no');
"#,
        &["1|1|no"],
    );
}
