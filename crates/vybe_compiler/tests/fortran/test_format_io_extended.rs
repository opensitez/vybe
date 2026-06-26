//! Extended Fortran formatted print/write coverage: I, F, A, E, L edit
//! descriptors, multi-value formats, parenthesized specs, labeled FORMAT,
//! and width/precision variants.

fortran_cases! {
    // ── I edit descriptor ────────────────────────────────────────────

    fmt_i5_positive_integer => {
        "program t\nprint '(I5)', 42\nend program t\n",
        ["42"]
    };

    fmt_i0_minimal_width => {
        "program t\nprint '(I0)', 12345\nend program t\n",
        ["12345"]
    };

    fmt_i10_ignored_field_width => {
        "program t\nprint '(I10)', 7\nend program t\n",
        ["7"]
    };

    fmt_i_negative_value => {
        "program t\nprint '(I5)', -99\nend program t\n",
        ["-99"]
    };

    fmt_i_zero => {
        "program t\nprint '(I4)', 0\nend program t\n",
        ["0"]
    };

    fmt_i_large_integer => {
        "program t\nprint '(I12)', 987654321\nend program t\n",
        ["987654321"]
    };

    fmt_i_variable_operand => {
        "program t\ninteger :: n = 17\nprint '(I0)', n\nend program t\n",
        ["17"]
    };

    fmt_i_expression_operand => {
        "program t\nprint '(I0)', 10 + 5\nend program t\n",
        ["15"]
    };

    // ── F edit descriptor ────────────────────────────────────────────

    fmt_f82_pi => {
        "program t\nprint '(F8.2)', 3.14159\nend program t\n",
        ["3.14"]
    };

    fmt_f104_pi => {
        "program t\nprint '(F10.4)', 3.14159\nend program t\n",
        ["3.1416"]
    };

    fmt_f62_zero => {
        "program t\nprint '(F6.2)', 0.0\nend program t\n",
        ["0.00"]
    };

    fmt_f51_half => {
        "program t\nprint '(F5.1)', 0.5\nend program t\n",
        ["0.5"]
    };

    fmt_f63_negative => {
        "program t\nprint '(F6.3)', -2.5\nend program t\n",
        ["-2.500"]
    };

    fmt_f_variable_real => {
        "program t\nreal :: x = 1.25\nprint '(F6.2)', x\nend program t\n",
        ["1.25"]
    };

    fmt_f_double_precision_literal => {
        "program t\nprint '(F10.4)', 2.718281828d0\nend program t\n",
        ["2.7183"]
    };

    // ── A edit descriptor ────────────────────────────────────────────

    fmt_a_plain_string => {
        "program t\nprint '(A)', 'hello'\nend program t\n",
        ["hello"]
    };

    fmt_a_empty_string => {
        "program t\nprint '(A)', ''\nend program t\n",
        [""]
    };

    fmt_a_character_variable => {
        "program t\ncharacter(len=5) :: s = 'world'\nprint '(A)', s\nend program t\n",
        ["world"]
    };

    fmt_a_with_literal_prefix => {
        "program t\nprint '(A, I0)', 'count=', 8\nend program t\n",
        ["count=8"]
    };

    fmt_a_two_strings => {
        "program t\nprint '(2A)', 'ab', 'cd'\nend program t\n",
        ["abcd"]
    };

    // ── E edit descriptor ────────────────────────────────────────────

    fmt_e124_large => {
        "program t\nprint '(E12.4)', 1.23456e10\nend program t\n",
        ["1.2346e+10"]
    };

    fmt_e102_small => {
        "program t\nprint '(E10.2)', 0.001\nend program t\n",
        ["1.00e-3"]
    };

    fmt_es103_quarter => {
        "program t\nprint '(ES10.3)', 0.25\nend program t\n",
        ["2.500e-1"]
    };

    fmt_e_negative_exponent => {
        "program t\nprint '(E11.3)', -4.5e3\nend program t\n",
        ["-4.500e+3"]
    };

    // ── L edit descriptor ────────────────────────────────────────────

    fmt_l5_true => {
        "program t\nprint '(L5)', .true.\nend program t\n",
        ["true"]
    };

    fmt_l5_false => {
        "program t\nprint '(L5)', .false.\nend program t\n",
        ["false"]
    };

    fmt_l_logical_variable => {
        "program t\nlogical :: ok = .true.\nprint '(L5)', ok\nend program t\n",
        ["true"]
    };

    fmt_2l5_both_values => {
        "program t\nprint '(2L5)', .true., .false.\nend program t\n",
        ["true false"]
    };

    // ── Multiple values in one format ────────────────────────────────

    fmt_multi_i_and_f => {
        "program t\ninteger :: n = 7\nreal :: x = 2.5\nprint '(I0, F6.2)', n, x\nend program t\n",
        ["7 2.50"]
    };

    fmt_multi_a_i_l => {
        "program t\nprint '(A, I0, L5)', 'flag=', 1, .true.\nend program t\n",
        ["flag=1true"]
    };

    fmt_multi_three_integers => {
        "program t\nprint '(3I0)', 1, 2, 3\nend program t\n",
        ["1 2 3"]
    };

    fmt_multi_two_reals => {
        "program t\nprint '(2F6.2)', 1.5, 2.25\nend program t\n",
        ["1.50 2.25"]
    };

    fmt_multi_mixed_four => {
        "program t\nprint '(A, I0, F5.1, L5)', 'v', 3, 1.4, .false.\nend program t\n",
        ["v3 1.4false"]
    };

    fmt_repeat_i4_descriptor => {
        "program t\nprint '(3I4)', 10, 20, 30\nend program t\n",
        ["10 20 30"]
    };

    // ── Parenthesized format strings ─────────────────────────────────

    fmt_paren_double_quoted => {
        "program t\nprint \"(I5)\", 42\nend program t\n",
        ["42"]
    };

    fmt_paren_literal_and_i => {
        "program t\nprint '(\"x=\", I0)', 5\nend program t\n",
        ["x=5"]
    };

    fmt_paren_nested_parens_in_literal => {
        "program t\nprint '(A, I0)', '(', 99\nend program t\n",
        ["(99"]
    };

    fmt_write_paren_format => {
        "program t\nwrite(*, '(I0)') 88\nend program t\n",
        ["88"]
    };

    fmt_print_apostrophe_format => {
        "program t\nprint '(A)', 'formatted'\nend program t\n",
        ["formatted"]
    };

    // ── Labeled FORMAT statements ────────────────────────────────────

    fmt_label_write_integer => {
        "program t\ninteger :: i = 7\nwrite(*, 100) i\n100 format(I5)\nend program t\n",
        ["7"]
    };

    fmt_label_write_real => {
        "program t\nreal :: x = 2.718\nwrite(*, 200) x\n200 format(F8.3)\nend program t\n",
        ["2.718"]
    };

    fmt_label_write_string => {
        "program t\ncharacter(len=5) :: s = 'hello'\nwrite(*, 300) s\n300 format(A)\nend program t\n",
        ["hello"]
    };

    fmt_label_write_logical => {
        "program t\nlogical :: flag = .false.\nwrite(*, 400) flag\n400 format(L5)\nend program t\n",
        ["false"]
    };

    fmt_label_write_multi => {
        "program t\ninteger :: n = 3\nreal :: r = 1.5\nwrite(*, 500) n, r\n500 format(I0, F5.1)\nend program t\n",
        ["3 1.5"]
    };

    // ── Width, precision, and spacing ────────────────────────────────

    fmt_f_precision_zero => {
        "program t\nprint '(F4.0)', 9.6\nend program t\n",
        ["10"]
    };

    fmt_f_precision_one => {
        "program t\nprint '(F5.1)', 9.64\nend program t\n",
        ["9.6"]
    };

    fmt_e_precision_six => {
        "program t\nprint '(E14.6)', 1.0\nend program t\n",
        ["1.000000e+0"]
    };

    fmt_5x_spacing => {
        "program t\nprint '(A, 5X, A)', 'L', 'R'\nend program t\n",
        ["L     R"]
    };

    fmt_a_e_combined_spacing => {
        "program t\nprint '(A, E10.2)', 'val=', 3.14\nend program t\n",
        ["val= 3.14e+0"]
    };

    fmt_i_f_sign_field_gap => {
        "program t\nprint '(I0, F6.2)', 1, 2.0\nend program t\n",
        ["1 2.00"]
    };

    fmt_three_f_different_precision => {
        "program t\nprint '(F4.1, F5.2, F6.3)', 1.1, 2.22, 3.333\nend program t\n",
        ["1.1 2.22 3.333"]
    };
}
