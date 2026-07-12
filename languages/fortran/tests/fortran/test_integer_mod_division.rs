//! Fortran integer mod(), modulo(), and truncating division semantics.

fortran_cases! {
    mod_positive_dividend_positive_divisor_23_7 => {
        "program t\nprint *, mod(23, 7)\nend program t\n",
        ["2"]
    };

    mod_positive_dividend_positive_divisor_31_8 => {
        "program t\nprint *, mod(31, 8)\nend program t\n",
        ["7"]
    };

    mod_positive_dividend_positive_divisor_14_6 => {
        "program t\nprint *, mod(14, 6)\nend program t\n",
        ["2"]
    };

    mod_negative_dividend_positive_divisor_neg17_5 => {
        "program t\nprint *, mod(-17, 5)\nend program t\n",
        ["-2"]
    };

    mod_negative_dividend_positive_divisor_neg23_7 => {
        "program t\nprint *, mod(-23, 7)\nend program t\n",
        ["-2"]
    };

    mod_negative_dividend_positive_divisor_neg11_4 => {
        "program t\nprint *, mod(-11, 4)\nend program t\n",
        ["-3"]
    };

    mod_negative_dividend_positive_divisor_neg1_5 => {
        "program t\nprint *, mod(-1, 5)\nend program t\n",
        ["-1"]
    };

    mod_positive_dividend_negative_divisor_17_neg5 => {
        "program t\nprint *, mod(17, -5)\nend program t\n",
        ["2"]
    };

    mod_positive_dividend_negative_divisor_23_neg7 => {
        "program t\nprint *, mod(23, -7)\nend program t\n",
        ["2"]
    };

    mod_positive_dividend_negative_divisor_11_neg4 => {
        "program t\nprint *, mod(11, -4)\nend program t\n",
        ["3"]
    };

    mod_both_negative_dividend_and_divisor_neg17_neg5 => {
        "program t\nprint *, mod(-17, -5)\nend program t\n",
        ["-2"]
    };

    mod_both_negative_dividend_and_divisor_neg23_neg7 => {
        "program t\nprint *, mod(-23, -7)\nend program t\n",
        ["-2"]
    };

    mod_both_negative_dividend_and_divisor_neg11_neg4 => {
        "program t\nprint *, mod(-11, -4)\nend program t\n",
        ["-3"]
    };

    modulo_positive_dividend_positive_divisor_23_7 => {
        "program t\nprint *, modulo(23, 7)\nend program t\n",
        ["2"]
    };

    modulo_negative_dividend_positive_divisor_neg10_3 => {
        "program t\nprint *, modulo(-10, 3)\nend program t\n",
        ["2"]
    };

    modulo_negative_dividend_positive_divisor_neg17_5 => {
        "program t\nprint *, modulo(-17, 5)\nend program t\n",
        ["3"]
    };

    modulo_positive_dividend_negative_divisor_10_neg3 => {
        "program t\nprint *, modulo(10, -3)\nend program t\n",
        ["-2"]
    };

    modulo_positive_dividend_negative_divisor_17_neg5 => {
        "program t\nprint *, modulo(17, -5)\nend program t\n",
        ["-3"]
    };

    modulo_both_negative_dividend_and_divisor_neg10_neg3 => {
        "program t\nprint *, modulo(-10, -3)\nend program t\n",
        ["-1"]
    };

    modulo_both_negative_dividend_and_divisor_neg17_neg5 => {
        "program t\nprint *, modulo(-17, -5)\nend program t\n",
        ["-2"]
    };

    mod_vs_modulo_negative_dividend_positive_divisor_neg7_3 => {
        "program t\nprint *, mod(-7, 3)\nprint *, modulo(-7, 3)\nend program t\n",
        ["-1", "2"]
    };

    mod_vs_modulo_negative_dividend_positive_divisor_neg13_4 => {
        "program t\nprint *, mod(-13, 4)\nprint *, modulo(-13, 4)\nend program t\n",
        ["-1", "3"]
    };

    mod_vs_modulo_positive_dividend_negative_divisor_7_neg3 => {
        "program t\nprint *, mod(7, -3)\nprint *, modulo(7, -3)\nend program t\n",
        ["1", "-2"]
    };

    mod_vs_modulo_positive_dividend_negative_divisor_13_neg4 => {
        "program t\nprint *, mod(13, -4)\nprint *, modulo(13, -4)\nend program t\n",
        ["1", "-3"]
    };

    integer_division_truncates_toward_zero_17_div_5 => {
        "program t\nprint *, 17 / 5\nend program t\n",
        ["3"]
    };

    integer_division_truncates_toward_zero_neg17_div_5 => {
        "program t\nprint *, -17 / 5\nend program t\n",
        ["-3"]
    };

    integer_division_truncates_toward_zero_17_div_neg5 => {
        "program t\nprint *, 17 / -5\nend program t\n",
        ["-3"]
    };

    integer_division_truncates_toward_zero_neg17_div_neg5 => {
        "program t\nprint *, -17 / -5\nend program t\n",
        ["3"]
    };

    integer_division_truncates_toward_zero_7_div_2 => {
        "program t\nprint *, 7 / 2\nend program t\n",
        ["3"]
    };

    integer_division_truncates_toward_zero_neg7_div_2 => {
        "program t\nprint *, -7 / 2\nend program t\n",
        ["-3"]
    };

    integer_division_truncates_toward_zero_7_div_neg2 => {
        "program t\nprint *, 7 / -2\nend program t\n",
        ["-3"]
    };

    integer_division_truncates_toward_zero_neg7_div_neg2 => {
        "program t\nprint *, -7 / -2\nend program t\n",
        ["3"]
    };

    mod_zero_result_exact_multiple_20_5 => {
        "program t\nprint *, mod(20, 5)\nend program t\n",
        ["0"]
    };

    mod_zero_result_exact_multiple_36_9 => {
        "program t\nprint *, mod(36, 9)\nend program t\n",
        ["0"]
    };

    mod_zero_result_exact_multiple_neg24_6 => {
        "program t\nprint *, mod(-24, 6)\nend program t\n",
        ["0"]
    };

    mod_zero_result_exact_multiple_24_neg6 => {
        "program t\nprint *, mod(24, -6)\nend program t\n",
        ["0"]
    };

    modulo_zero_result_exact_multiple_20_5 => {
        "program t\nprint *, modulo(20, 5)\nend program t\n",
        ["0"]
    };

    modulo_zero_result_exact_multiple_neg24_6 => {
        "program t\nprint *, modulo(-24, 6)\nend program t\n",
        ["0"]
    };

    do_loop_counts_multiples_of_three_up_to_15 => {
        "program t\ninteger :: i, c\nc = 0\ndo i = 1, 15\nif (mod(i, 3) == 0) c = c + 1\nend do\nprint *, c\nend program t\n",
        ["5"]
    };

    do_loop_counts_even_numbers_up_to_20 => {
        "program t\ninteger :: i, c\nc = 0\ndo i = 1, 20\nif (mod(i, 2) == 0) c = c + 1\nend do\nprint *, c\nend program t\n",
        ["10"]
    };

    do_loop_counts_fizzbuzz_multiples_of_15 => {
        "program t\ninteger :: i, c\nc = 0\ndo i = 1, 100\nif (mod(i, 15) == 0) c = c + 1\nend do\nprint *, c\nend program t\n",
        ["6"]
    };

    do_loop_stops_when_mod_reaches_zero => {
        "program t\ninteger :: i, c\nc = 0\ndo i = 1, 100\nif (mod(i, 7) == 0) c = c + 1\nif (c == 5) exit\nend do\nprint *, c\nend program t\n",
        ["5"]
    };

    do_while_mod_condition_counts_steps => {
        "program t\ninteger :: n, steps\nn = 100\nsteps = 0\ndo while (n > 1)\nn = n / 2\nsteps = steps + 1\nend do\nprint *, steps\nend program t\n",
        ["6"]
    };

    gcd_loop_mod_35_and_14 => {
        "program t\ninteger :: a, b, tmp\na = 35\nb = 14\ndo while (b /= 0)\ntmp = b\nb = mod(a, b)\na = tmp\nend do\nprint *, a\nend program t\n",
        ["7"]
    };

    gcd_loop_mod_1071_and_462 => {
        "program t\ninteger :: a, b, tmp\na = 1071\nb = 462\ndo while (b /= 0)\ntmp = b\nb = mod(a, b)\na = tmp\nend do\nprint *, a\nend program t\n",
        ["21"]
    };

    gcd_loop_mod_270_and_192 => {
        "program t\ninteger :: a, b, tmp\na = 270\nb = 192\ndo while (b /= 0)\ntmp = b\nb = mod(a, b)\na = tmp\nend do\nprint *, a\nend program t\n",
        ["6"]
    };

    gcd_loop_mod_99_and_78 => {
        "program t\ninteger :: a, b, tmp\na = 99\nb = 78\ndo while (b /= 0)\ntmp = b\nb = mod(a, b)\na = tmp\nend do\nprint *, a\nend program t\n",
        ["3"]
    };

    mod_with_variables_positive_case => {
        "program t\ninteger :: a = 29, b = 6\nprint *, mod(a, b)\nend program t\n",
        ["5"]
    };

    mod_with_variables_negative_dividend => {
        "program t\ninteger :: a = -29, b = 6\nprint *, mod(a, b)\nend program t\n",
        ["-5"]
    };

    integer_division_with_variables_truncates => {
        "program t\ninteger :: a = -29, b = 6\nprint *, a / b\nend program t\n",
        ["-4"]
    };

    mod_reconstructs_dividend_from_quotient => {
        "program t\ninteger :: a = 29, b = 6, q, r\nq = a / b\nr = mod(a, b)\nprint *, q * b + r\nend program t\n",
        ["29"]
    };

    modulo_reconstructs_negative_dividend => {
        "program t\ninteger :: a = -29, b = 6, q, r\nq = a / b\nr = modulo(a, b)\nprint *, q * b + r\nend program t\n",
        ["-29"]
    };
}
