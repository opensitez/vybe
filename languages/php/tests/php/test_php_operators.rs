use super::helpers::run_prints;

fn assert_output(expr: &str, expected: &str) {
    assert_eq!(run_prints(&format!("<?php echo {}; ", expr)), vec![expected.to_string()]);
}

fn assert_int(expr: &str, expected: i64) {
    assert_output(expr, &expected.to_string());
}

fn assert_bool(expr: &str, expected: bool) {
    assert_output(expr, if expected { "1" } else { "0" });
}

#[test]
fn php_operator_integer_arithmetic() {
    let values: [i64; 3] = [0, 1, 6];

    for a in values {
        for b in values {
            assert_int(&format!("({a} + {b})"), a + b);
            assert_int(&format!("({a} - {b})"), a - b);
            assert_int(&format!("({a} * {b})"), a * b);

            if b != 0 {
                assert_int(&format!("({a} % {b})"), a % b);
                assert_int(&format!("intdiv({a}, {b})"), a / b);
                if a % b == 0 {
                    assert_int(&format!("({a} / {b})"), a / b);
                }
            }
        }
    }

    for base in [0_i64, 1, 7] {
        for exp in [0_i64, 1, 4] {
            assert_int(&format!("({base} ** {exp})"), base.pow(exp as u32));
        }
    }

    for base in [0_i64, 1, 9] {
        for left_shift in [0_i64, 1, 5] {
            assert_int(&format!("({base} << {left_shift})"), base << left_shift);
            assert_int(&format!("({base} >> {left_shift})"), base >> left_shift);
            assert_int(&format!("({base} & (1 << {left_shift}))"), base & (1_i64 << left_shift));
            assert_int(&format!("({base} | (1 << {left_shift}))"), base | (1_i64 << left_shift));
            assert_int(&format!("({base} ^ (1 << {left_shift}))"), base ^ (1_i64 << left_shift));
        }
    }
}

#[test]
fn php_operator_comparison() {
    let values: [i64; 3] = [-3, 0, 3];

    for a in values {
        for b in values {
            assert_bool(&format!("({a} == {b})"), a == b);
            assert_bool(&format!("({a} != {b})"), a != b);
            assert_bool(&format!("({a} < {b})"), a < b);
            assert_bool(&format!("({a} > {b})"), a > b);
            assert_bool(&format!("({a} <= {b})"), a <= b);
            assert_bool(&format!("({a} >= {b})"), a >= b);
            assert_int(&format!("({a} <=> {b})"), (a > b) as i8 as i64 - (a < b) as i8 as i64);
        }
    }

    let nullish = [
        ("NULL", "'fallback'", "fallback"),
        ("'left'", "'fallback'", "left"),
        ("0", "'fallback'", "0"),
        ("1", "'fallback'", "1"),
        ("'0'", "'fallback'", "0"),
    ];

    for (left_expr, right_expr, expected) in nullish {
        assert_output(&format!("({left_expr} ?? {right_expr})"), expected);
    }
}

#[test]
fn php_operator_ternary_truthiness_and_string_ops() {
    for value in 0_i64..8_i64 {
        let expected = if value % 2 == 0 { "even" } else { "odd" };
        assert_output(
            &format!("(({value} % 2) == 0 ? 'even' : 'odd')"),
            expected,
        );
    }

    let left_values = ["true", "false", "true", "false"];
    let right_values = ["true", "true", "false", "false"];

    for i in 0..left_values.len() {
        let a = left_values[i];
        let b = right_values[i];
        assert_bool(&format!("({a} and {b})"), (a == "true") && (b == "true"));
        assert_bool(&format!("({a} or {b})"), (a == "true") || (b == "true"));
        assert_bool(&format!("({a} xor {b})"), (a == "true") ^ (b == "true"));
    }

    let prefixes = ["App", "Kernel", "Http", "Service", "Domain"]; 
    let suffixes = ["Model", "Repository", "Controller", "Request", "Response"]; 

    for prefix in prefixes {
        for suffix in suffixes {
            let expected = format!("{prefix}{suffix}");
            assert_output(
                &format!("'{prefix}' . '{suffix}'"),
                &expected,
            );
        }
    }
}

#[test]
fn php_operator_identity_and_coercion_edges() {
    assert_bool("(1 == '1')", true);
    assert_bool("(1 === '1')", false);
    assert_bool("(1 != '1')", false);
    assert_bool("(1 !== '1')", true);
    assert_bool("('' == false)", true);
    assert_bool("('' === false)", false);
    assert_bool("('' == 0)", true);
    assert_bool("('' === 0)", false);
    assert_bool("('0' == 0)", true);
    assert_bool("('0' == false)", true);
    assert_bool("('0' === false)", false);
    assert_bool("(0 == false)", true);
    assert_bool("(0 === false)", false);
    assert_bool("(null == false)", true);
    assert_bool("(null === false)", false);
    assert_bool("(null == [])", true);
    assert_bool("(null === [])", false);
}

#[test]
fn php_operator_truthiness_explicit_casts() {
    assert_bool("(bool) null", false);
    assert_bool("(bool) 0", false);
    assert_bool("(bool) 1", true);
    assert_bool("(bool) -1", true);
    assert_bool("(bool) 0.0", false);
    assert_bool("(bool) 0.1", true);
    assert_bool("(bool) ''", false);
    assert_bool("(bool) '0'", false);
    assert_bool("(bool) '1'", true);
    assert_bool("(bool) []", false);
    assert_bool("(bool) [0]", true);
}

#[test]
fn php_operator_precedence_arithmetic_logical() {
    assert_int("(2 + 3 << 1)", (2 + (3 << 1)) as i64);
    assert_int("((2 + 3) << 1)", 10);
    assert_int("(2 << 1 + 1)", 8);
    assert_int("(2 << (1 + 1))", 8);
    assert_int("((5 + 3) > 6 && 3 < 10)", 1);
    assert_int("((5 + 3) > 2 && 1 === 1)", 1);
    assert_int("((5 + 3) > 2 ? 7 : 3)", 7);
    assert_int("((5 > 2 && 3 < 2) ? 1 : 0)", 0);
}

#[test]
fn php_operator_bitwise_and_shift_patterns() {
    assert_int("((3 | 1) & 2)", 2);
    assert_int("((7 ^ 3) >> 1)", 2);
    assert_int("((12 << 1) | 1)", 25);
    assert_int("(~0 & 7)", 7);
    assert_int("(8 & 15)", 8);
    assert_int("(8 | 1)", 9);
    assert_int("(12 ^ 5)", 9);
}

#[test]
fn php_operator_logical_keyword_symbol_parity() {
    assert_bool("(true and false)", false);
    assert_bool("(true && false)", false);
    assert_bool("(true or false)", true);
    assert_bool("(true || false)", true);
    assert_bool("(true xor false)", true);
    assert_bool("(false xor false)", false);
    assert_bool("(!true)", false);
    assert_bool("(!false)", true);
    assert_bool("(!!false)", false);
    assert_bool("(!('a' === 'a'))", false);
    assert_bool("(!('a' === 'b'))", true);
}

#[test]
fn php_operator_nullish_and_ternary_precedence_edges() {
    assert_output("(null ?? 'fallback')", "fallback");
    assert_output("('' ?? 'fallback')", "");
    assert_output("(0 ?? 'fallback')", "0");
    assert_output("(false ?? 'fallback')", "");

    assert_output("('' ?: 'fallback')", "fallback");
    assert_output("(0 ?: 'fallback')", "fallback");
    assert_output("(1 ?: 'fallback')", "1");

    assert_output("(1 + 2 ?: 0)", "3");
    assert_output("(0 ?: 1 + 2)", "3");

    assert_output("(true ? 'yes' : (false ? 'no' : 'late'))", "yes");
    assert_output("(false ? 'yes' : (false ? 'no' : 'late'))", "late");
    assert_output("(false ? 'a' : false ? 'b' : 'c')", "c");
}

#[test]
fn php_operator_right_associative_examples() {
    assert_output("(1 <=> 2) > 0", "0");
    assert_output("(2 <=> 1) > 0", "1");
    assert_output("((1 <=> 1) === 0)", "1");
    assert_output("('1' <=> '2')", "-1");
    assert_output("('a' <=> 'b')", "-1");
    assert_output("(true <=> false)", "1");
    assert_output("(false <=> true)", "-1");
}

#[test]
fn php_operator_array_equality_structures() {
    assert_bool("([1,2] == [1,2])", true);
    assert_bool("([1,2] === [1,2])", true);
    assert_bool("([1,2] == [2,1])", false);
    assert_bool("([1,'2'] == [1,2])", true);
    assert_bool("([1,'2'] === [1,2])", false);
    assert_bool("(['a' => 1] == ['a' => 1.0])", true);
    assert_bool("(['a' => 1] === ['a' => 1.0])", false);
}

#[test]
fn php_operator_match_like_condition_shape() {
    assert_output(
        "match (5) { 1 => 'one', 2, 3 => 'small', 4, 5 => 'mid', default => 'other' }",
        "mid",
    );
    assert_output(
        "match (true) { 1 > 2 => 'bad', 2 + 2 === 5 => 'bad', default => 'good' }",
        "good",
    );
    assert_output(
        "match (false) { true => 'no', false => 'yes', default => 'bad' }",
        "yes",
    );
}

#[test]
fn php_operator_precedence_and_parentheses() {
    assert_int("(1 + 2 * 3)", 7);
    assert_int("((1 + 2) * 3)", 9);
    assert_int("(10 / 2 + 3 * 4)", 17);
    assert_int("(10 / (2 + 3) * 4)", 8);
    assert_int("(1 + 2 - 3 + 4 * 5)", 20);
    assert_int("((1 + 2 - 3) + 4 * 5)", 18);
    assert_int("((2 ** 3) ** 2)", 64);
    assert_int("(2 ** 3 ** 2)", 512);
    assert_int("(-2) ** 2", 4);
    assert_int("-(2 ** 2)", -4);
}

#[test]
fn php_operator_boolean_keyword_precedence() {
    assert_bool(
        "(1 || 0 && 0)",
        true,
    );
    assert_bool(
        "((1 || 0) && 0)",
        false,
    );
    assert_output(
        "(function() { $a = true; $a = $a or false; return $a ? 1 : 0; })()",
        "1",
    );
    assert_output(
        "(function() { $a = true; $a = $a || false; return $a ? 1 : 0; })()",
        "0",
    );
    assert_output(
        "(function() { $a = false; $a = $a and true; return $a ? 1 : 0; })()",
        "0",
    );
    assert_output(
        "(function() { $a = false; $a = $a && true; return $a ? 1 : 0; })()",
        "0",
    );
}

#[test]
fn php_operator_compound_assignment() {
    assert_int(
        "(function() { $value = 10; $value += 5; $value -= 3; $value *= 2; $value /= 4; return $value; })()",
        6,
    );
    assert_int(
        "(function() { $value = 2; $value **= 3; return $value; })()",
        8,
    );
    assert_int(
        "(function() { $value = 9; $value %= 5; return $value; })()",
        4,
    );
    assert_int(
        "(function() { $value = 1; $value <<= 3; $value >>= 2; return $value; })()",
        2,
    );
    assert_int(
        "(function() { $value = 1; $value |= 2; $value &= 3; return $value; })()",
        2,
    );
    assert_int(
        "(function() { $value = 15; $value ^= 10; return $value; })()",
        5,
    );
}

#[test]
fn php_operator_coercion_edges() {
    assert_int("('10' == 10)", 1);
    assert_bool("('10' === 10)", false);
    assert_int("('01' == 1)", 1);
    assert_bool("('01' === 1)", false);
    assert_int("('' == 0)", 1);
    assert_bool("('' === 0)", false);
    assert_int("('0' == false)", 1);
    assert_bool("('0' === false)", false);
    assert_int("('0.0' == 0.0)", 1);
    assert_bool("('0.0' === 0.0)", false);
}

#[test]
fn php_operator_string_and_numeric_mix() {
    assert_output("('a' . 1 . true)", "a11");
    assert_output("('a' . (1 + true))", "a2");
    assert_output("(1 + '2')", "3");
    assert_output("('3' + '4')", "7");
    assert_output("('2' . ('1' + 2))", "23");
    assert_output("('value=' . (1 ? 2 : 3))", "value=2");
    assert_output("(function() { $left = 'left'; $right = null; return $left . ($right ?? 'fallback'); })()", "leftfallback");
    assert_output("(function() { $left = 'left'; $right = 'right'; return $left . ($right ?? 'fallback'); })()", "leftright");
}

#[test]
fn php_operator_null_coalescing_and_safe_navigation_edges() {
    assert_output(
        "(function() { $value = null; $value ??= 'default'; return $value; })()",
        "default",
    );
    assert_output(
        "(function() { $value = 0; $value ??= 'default'; return $value; })()",
        "0",
    );
    assert_output(
        "(function() { $user = null; return $user?->name ?? 'anon'; })()",
        "anon",
    );
    assert_output(
        "(function() { $user = (object)['name' => 'Ada']; return $user?->name ?? 'anon'; })()",
        "Ada",
    );
}

#[test]
fn php_operator_spaceship_truthiness_edges() {
    assert_int("(4 <=> 4)", 0);
    assert_int("(-1 <=> 2)", -1);
    assert_int("(2 <=> -1)", 1);
    assert_int("(true <=> false)", 1);
    assert_int("(false <=> true)", -1);
    assert_output("(null <=> null)", "0");
    assert_output("(true ?: 'fallback')", "1");
    assert_output("((true && false) ?: 'fallback')", "fallback");
    assert_output("((0 ?: 1) <=> (1 ?: 0))", "0");
}
