//! Fortran expression precedence and associativity: exponentiation, arithmetic,
//! logical (.and./.or.), parenthesis overrides, and mixed int/real promotion.

fortran_cases! {
    // ── ** right-associativity ────────────────────────────────────────

    power_right_assoc_two_cubed_squared => {
        "program t\nprint *, 2 ** 3 ** 2\nend program t\n",
        ["512"]
    };

    power_right_assoc_five_squared_cubed => {
        "program t\nprint *, 5 ** 2 ** 3\nend program t\n",
        ["390625"]
    };

    power_before_multiply_three_two_cubed => {
        "program t\nprint *, 3 * 2 ** 3\nend program t\n",
        ["24"]
    };

    // ── * / vs + - ───────────────────────────────────────────────────

    multiply_before_add_three_four_five => {
        "program t\nprint *, 3 + 4 * 5\nend program t\n",
        ["23"]
    };

    multiply_before_add_six_two_plus_one => {
        "program t\nprint *, 6 * 2 + 1\nend program t\n",
        ["13"]
    };

    divide_before_add_twelve_over_three => {
        "program t\nprint *, 12 / 3 + 1\nend program t\n",
        ["5"]
    };

    subtract_left_assoc_ten_four_two => {
        "program t\nprint *, 10 - 4 - 2\nend program t\n",
        ["4"]
    };

    divide_multiply_left_assoc_eight => {
        "program t\nprint *, 8 / 2 * 3\nend program t\n",
        ["12"]
    };

    // ── unary minus vs ** ────────────────────────────────────────────

    unary_minus_after_power_two_fourth => {
        "program t\nprint *, -2 ** 4\nend program t\n",
        ["-16"]
    };

    unary_minus_after_power_three_squared => {
        "program t\nprint *, -3 ** 2\nend program t\n",
        ["-9"]
    };

    unary_minus_power_then_add => {
        "program t\nprint *, -2 ** 3 + 1\nend program t\n",
        ["-7"]
    };

    // ── .and. vs .or. ────────────────────────────────────────────────

    and_binds_tighter_than_or_true_or_false_and => {
        "program t\nprint *, .true. .or. .false. .and. .false.\nend program t\n",
        ["true"]
    };

    and_binds_tighter_than_or_false_and_true_or => {
        "program t\nprint *, .false. .and. .true. .or. .true.\nend program t\n",
        ["true"]
    };

    and_binds_tighter_than_or_false_or_true_and => {
        "program t\nprint *, .false. .or. .true. .and. .false.\nend program t\n",
        ["false"]
    };

    paren_or_before_and => {
        "program t\nprint *, (.false. .or. .true.) .and. .false.\nend program t\n",
        ["false"]
    };

    // ── parenthesis overrides ────────────────────────────────────────

    paren_sum_before_square => {
        "program t\nprint *, (1 + 2) ** 2\nend program t\n",
        ["9"]
    };

    paren_sum_times_paren_sum => {
        "program t\nprint *, (2 + 3) * (4 + 1)\nend program t\n",
        ["25"]
    };

    paren_negated_power_cube => {
        "program t\nprint *, -(2 ** 3)\nend program t\n",
        ["-8"]
    };

    paren_quotient_of_differences => {
        "program t\nprint *, (10 - 2) / (3 - 1)\nend program t\n",
        ["4"]
    };

    // ── mixed int/real in precedence chains ──────────────────────────

    int_real_mult_in_add_chain => {
        "program t\nprint *, 1 + 2 * 3.0\nend program t\n",
        ["7"]
    };

    power_then_real_add => {
        "program t\nprint *, 2 ** 2 + 1.0\nend program t\n",
        ["5"]
    };

    int_real_sub_mult_chain => {
        "program t\nprint *, 10 - 3 * 2.0\nend program t\n",
        ["4"]
    };

    real_divide_then_add => {
        "program t\nprint *, 4 / 2 + 1.0\nend program t\n",
        ["3"]
    };

    not_binds_before_and => {
        "program t\nprint *, .not. .false. .and. .false.\nend program t\n",
        ["false"]
    };

    unary_plus_is_neutral => {
        "program t\nprint *, +(-2)\nend program t\n",
        ["-2"]
    };

    comparison_before_logical_and => {
        "program t\nprint *, 1 + 2 == 3 .and. .true.\nend program t\n",
        ["true"]
    };

    negative_division_with_parentheses => {
        "program t\nprint *, -(7 / 2)\nprint *, (7 / 2)\nend program t\n",
        ["-3", "3"]
    };
}
