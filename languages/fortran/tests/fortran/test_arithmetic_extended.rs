//! Extended Fortran arithmetic: exponentiation, unary +/-, mixed real/integer
//! promotion, parenthesis precedence, chained operations, integer overflow wrap,
//! real division, and comparison expressions printing 0/1.

fortran_cases! {
    // ── Exponentiation ───────────────────────────────────────────────

    power_zero_exponent_is_one => {
        "program t\nprint *, 5 ** 0\nend program t\n",
        ["1"]
    };

    power_two_to_eight => {
        "program t\nprint *, 2 ** 8\nend program t\n",
        ["256"]
    };

    power_negative_base_odd_exponent => {
        "program t\nprint *, (-2) ** 3\nend program t\n",
        ["-8"]
    };

    power_negative_base_even_exponent => {
        "program t\nprint *, (-2) ** 4\nend program t\n",
        ["16"]
    };

    power_right_associative_four_cubed_squared => {
        "program t\nprint *, 4 ** (3 ** 2)\nend program t\n",
        ["262144"]
    };

    power_right_associative_three_squared_squared => {
        "program t\nprint *, 3 ** (2 ** 2)\nend program t\n",
        ["81"]
    };

    power_before_addition => {
        "program t\nprint *, 1 + 2 ** 4\nend program t\n",
        ["17"]
    };

    power_of_parenthesized_sum => {
        "program t\nprint *, (2 + 1) ** 3\nend program t\n",
        ["27"]
    };

    power_times_multiplier => {
        "program t\nprint *, 2 ** 3 * 4\nend program t\n",
        ["32"]
    };

    // ── Unary plus and minus ─────────────────────────────────────────

    unary_plus_on_variable => {
        "program t\ninteger :: x = 42\nprint *, +x\nend program t\n",
        ["42"]
    };

    unary_minus_on_literal => {
        "program t\nprint *, -17\nend program t\n",
        ["-17"]
    };

    unary_minus_on_parenthesized_sum => {
        "program t\nprint *, -(3 + 4)\nend program t\n",
        ["-7"]
    };

    double_unary_minus => {
        "program t\ninteger :: x = 5\nprint *, -(-x)\nend program t\n",
        ["5"]
    };

    unary_minus_on_real_literal => {
        "program t\nprint *, -3.5\nend program t\n",
        ["-3.5"]
    };

    // ── Mixed real/integer promotion ─────────────────────────────────

    integer_plus_real_literal => {
        "program t\nprint *, 2 + 3.0\nend program t\n",
        ["5"]
    };

    integer_times_real_literal => {
        "program t\nprint *, 5 * 2.0\nend program t\n",
        ["10"]
    };

    real_minus_integer_literal => {
        "program t\nprint *, 10.0 - 3\nend program t\n",
        ["7"]
    };

    real_divided_by_integer_literal => {
        "program t\nprint *, 6.0 / 2\nend program t\n",
        ["3"]
    };

    integer_divided_by_real_promotes => {
        "program t\nprint *, 7 / 2.0\nend program t\n",
        ["3.5"]
    };

    integer_power_real_exponent => {
        "program t\nprint *, 2 ** 3.0\nend program t\n",
        ["8"]
    };

    // ── Parenthesis precedence ───────────────────────────────────────

    subtraction_before_multiplication => {
        "program t\nprint *, 10 - 3 * 2\nend program t\n",
        ["4"]
    };

    division_before_addition => {
        "program t\nprint *, 20 / 4 + 3\nend program t\n",
        ["8"]
    };

    mixed_add_multiply_subtract => {
        "program t\nprint *, 2 + 3 * 4 - 5\nend program t\n",
        ["9"]
    };

    division_of_parenthesized_sum => {
        "program t\nprint *, 48 / (6 + 2)\nend program t\n",
        ["6"]
    };

    nested_parentheses_left_associative => {
        "program t\nprint *, ((1 + 2) * 3) + 4\nend program t\n",
        ["13"]
    };

    left_associative_subtraction_chain => {
        "program t\nprint *, 8 - 3 - 2\nend program t\n",
        ["3"]
    };

    left_associative_division_chain => {
        "program t\nprint *, 100 / 10 / 2\nend program t\n",
        ["5"]
    };

    // ── Chained operations ─────────────────────────────────────────────

    chained_add_multiply_assign => {
        "program t\ninteger :: x\nx = 1 + 2 ** 3 * 4\nprint *, x\nend program t\n",
        ["33"]
    };

    chained_real_arithmetic_assign => {
        "program t\nreal :: y\ny = 1.5 * 2.0 + 3.5\nprint *, y\nend program t\n",
        ["6.5"]
    };

    chained_subtract_divide_add => {
        "program t\nprint *, 30 - 12 / 3 + 1\nend program t\n",
        ["27"]
    };

    chained_power_add_multiply => {
        "program t\nprint *, 3 ** 2 + 2 * 5\nend program t\n",
        ["19"]
    };

    chained_unary_in_expression => {
        "program t\ninteger :: a = 3, b = 7\nprint *, a + -b\nend program t\n",
        ["-4"]
    };

    // ── Integer overflow wrap (visible bit/sign wrap) ──────────────────

    ishft_one_into_sign_bit_wraps => {
        "program t\nprint *, ishft(1, 31)\nend program t\n",
        ["-2147483648"]
    };

    ishft_half_max_left_wraps_sign => {
        "program t\nprint *, ishft(1073741824, 1)\nend program t\n",
        ["-2147483648"]
    };

    near_max_integer_add_reaches_huge => {
        "program t\ninteger :: x\nx = 2147483640 + 7\nprint *, x\nend program t\n",
        ["2147483647"]
    };

    large_product_near_integer_limit => {
        "program t\ninteger :: x\nx = 46340 * 46340\nprint *, x\nend program t\n",
        ["2147395600"]
    };

    // ── Real division ──────────────────────────────────────────────────

    real_division_fifteen_over_six => {
        "program t\nprint *, 15.0 / 6.0\nend program t\n",
        ["2.5"]
    };

    real_division_seven_over_two => {
        "program t\nprint *, 7.0 / 2.0\nend program t\n",
        ["3.5"]
    };

    real_division_one_quarter => {
        "program t\nprint *, 1.0 / 4.0\nend program t\n",
        ["0.25"]
    };

    real_division_negative_over_positive => {
        "program t\nprint *, -8.0 / 4.0\nend program t\n",
        ["-2"]
    };

    // ── Comparisons in expressions (if/else → 0/1) ───────────────────

    cmp_eq_true_prints_one => {
        "program t\ninteger :: r\nif (5 == 5) then\nr = 1\nelse\nr = 0\nend if\nprint *, r\nend program t\n",
        ["1"]
    };

    cmp_eq_false_prints_zero => {
        "program t\ninteger :: r\nif (5 == 3) then\nr = 1\nelse\nr = 0\nend if\nprint *, r\nend program t\n",
        ["0"]
    };

    cmp_gt_true_prints_one => {
        "program t\ninteger :: r\nif (7 > 3) then\nr = 1\nelse\nr = 0\nend if\nprint *, r\nend program t\n",
        ["1"]
    };

    cmp_lt_false_prints_zero => {
        "program t\ninteger :: r\nif (2 < 1) then\nr = 1\nelse\nr = 0\nend if\nprint *, r\nend program t\n",
        ["0"]
    };

    cmp_ge_true_prints_one => {
        "program t\ninteger :: r\nif (10 >= 10) then\nr = 1\nelse\nr = 0\nend if\nprint *, r\nend program t\n",
        ["1"]
    };

    cmp_le_false_prints_zero => {
        "program t\ninteger :: r\nif (4 <= 3) then\nr = 1\nelse\nr = 0\nend if\nprint *, r\nend program t\n",
        ["0"]
    };

    cmp_ne_equal_false_prints_zero => {
        "program t\ninteger :: r\nif (8 /= 8) then\nr = 1\nelse\nr = 0\nend if\nprint *, r\nend program t\n",
        ["0"]
    };

    cmp_ne_distinct_prints_one => {
        "program t\ninteger :: r\nif (8 /= 5) then\nr = 1\nelse\nr = 0\nend if\nprint *, r\nend program t\n",
        ["1"]
    };

    cmp_on_arithmetic_equality => {
        "program t\ninteger :: r\nif (3 + 4 == 7) then\nr = 1\nelse\nr = 0\nend if\nprint *, r\nend program t\n",
        ["1"]
    };

    cmp_on_chained_comparison => {
        "program t\ninteger :: r\nif ((2 * 3) > 5) then\nr = 1\nelse\nr = 0\nend if\nprint *, r\nend program t\n",
        ["1"]
    };

    chained_unary_plus_minus_chain => {
        "program t\nprint *, 10 - -5\nend program t\n",
        ["15"]
    };

    unary_minus_after_parentheses_with_real_multiplication => {
        "program t\nprint *, -(2.0 * 3.0) + 1\nend program t\n",
        ["-5"]
    };

    unary_minus_covers_parenthesized_power_operand => {
        "program t\nprint *, -(2 ** 3)\nend program t\n",
        ["-8"]
    };

    mixed_real_fractional_roundtrip => {
        "program t\nprint *, 7 / 2 + 0.5\nend program t\n",
        ["3.5"]
    };

    left_associative_additive_with_unary_chain => {
        "program t\nprint *, 20 - 5 + -3\nend program t\n",
        ["12"]
    };

    unary_plus_noop_preserves_variable_sign => {
        "program t\ninteger :: x\nx = -9\nprint *, +x\nend program t\n",
        ["-9"]
    };
}
