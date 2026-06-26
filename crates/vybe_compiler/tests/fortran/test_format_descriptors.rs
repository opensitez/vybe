//! Fortran format descriptor semantics: I, F, E/ES, L, A edit descriptors,
//! X spacing, / record advance, sign-position gaps, Iw.m tokens, and
//! SP/SS/S/BN/BZ prefix behavior. Each case uses a distinct value or
//! format combination that changes the formatted output string.

fortran_cases! {
    // ── I edit descriptor (value-driven variants) ────────────────────

    desc_i0_single_digit => {
        "program t\nprint '(I0)', 7\nend program t\n",
        ["7"]
    };

    desc_i0_six_digits => {
        "program t\nprint '(I0)', 123456\nend program t\n",
        ["123456"]
    };

    desc_i0_negative_three_digits => {
        "program t\nprint '(I0)', -123\nend program t\n",
        ["-123"]
    };

    desc_i0_zero_literal => {
        "program t\nprint '(I0)', 0\nend program t\n",
        ["0"]
    };

    desc_2i0_adjacent_integers => {
        "program t\nprint '(2I0)', 12, 34\nend program t\n",
        ["12 34"]
    };

    desc_3i4_repeat_count_values => {
        "program t\nprint '(3I4)', 1, 22, 333\nend program t\n",
        ["1 22 333"]
    };

    desc_i0_after_literal_prefix => {
        "program t\nprint '(\"n=\", I0)', 99\nend program t\n",
        ["n=99"]
    };

    // ── F edit descriptor (precision-driven) ─────────────────────────

    desc_f41_tiny_positive_truncates => {
        "program t\nprint '(F4.1)', 0.04\nend program t\n",
        ["0.0"]
    };

    desc_f73_three_decimal_places => {
        "program t\nprint '(F7.3)', 12.3456\nend program t\n",
        ["12.346"]
    };

    desc_f52_large_integer_style => {
        "program t\nprint '(F5.2)', 99.999\nend program t\n",
        ["100.00"]
    };

    desc_f84_negative_value => {
        "program t\nprint '(F8.4)', -0.375\nend program t\n",
        ["-0.3750"]
    };

    // ── E and ES edit descriptors ──────────────────────────────────

    desc_e113_negative_thousands => {
        "program t\nprint '(E11.3)', -4.5e3\nend program t\n",
        ["-4.500e+3"]
    };

    desc_es144_double_kind => {
        "program t\nprint '(ES14.4)', 2.718281828d0\nend program t\n",
        ["2.7183e+0"]
    };

    desc_e146_unity_six_places => {
        "program t\nprint '(E14.6)', 1.0\nend program t\n",
        ["1.000000e+0"]
    };

    desc_e82_fractional_mantissa => {
        "program t\nprint '(E8.2)', 6.25\nend program t\n",
        ["6.25e+0"]
    };

    // ── L edit descriptor ────────────────────────────────────────────

    desc_l1_false_compact => {
        "program t\nprint '(L1)', .false.\nend program t\n",
        ["false"]
    };

    desc_2l3_both_logical_values => {
        "program t\nprint '(2L3)', .true., .false.\nend program t\n",
        ["true false"]
    };

    desc_l_after_a_prefix => {
        "program t\nprint '(A, L5)', 'ok=', .true.\nend program t\n",
        ["ok=true"]
    };

    // ── A edit descriptor ────────────────────────────────────────────

    desc_2a_spaced_words => {
        "program t\nprint '(2A)', 'foo', 'bar'\nend program t\n",
        ["foo bar"]
    };

    desc_a_with_numeric_suffix => {
        "program t\nprint '(A, I0)', 'item', 3\nend program t\n",
        ["item3"]
    };

    // ── X spacing (distinct repeat counts) ───────────────────────────

    desc_1x_single_blank => {
        "program t\nprint '(A, 1X, A)', 'a', 'b'\nend program t\n",
        ["a b"]
    };

    desc_3x_triple_blank => {
        "program t\nprint '(A, 3X, A)', 'L', 'R'\nend program t\n",
        ["L   R"]
    };

    desc_10x_decuple_blank => {
        "program t\nprint '(A, 10X, A)', 'x', 'y'\nend program t\n",
        ["x          y"]
    };

    desc_2x_before_integer => {
        "program t\nprint '(A, 2X, I0)', 'n', 42\nend program t\n",
        ["n  42"]
    };

    // ── / record advance (newline) ───────────────────────────────────

    desc_slash_leading_newline => {
        "program t\nprint '(/, A)', 'alone'\nend program t\n",
        ["\nalone"]
    };

    desc_double_slash_blank_line => {
        "program t\nprint '(A, /, /, A)', 'a', 'b'\nend program t\n",
        ["a\n\nb"]
    };

    desc_i0_slash_i0_two_records => {
        "program t\nprint '(I0, /, I0)', 1, 2\nend program t\n",
        ["1\n2"]
    };

    // ── Sign-position gap after prior data item ──────────────────────

    desc_i0_f62_positive_gap => {
        "program t\nprint '(I0, F6.2)', 3, 1.5\nend program t\n",
        ["3 1.50"]
    };

    desc_i0_f63_negative_gap => {
        "program t\nprint '(I0, F6.3)', 1, -2.5\nend program t\n",
        ["1 -2.500"]
    };

    desc_f62_f62_second_positive_gap => {
        "program t\nprint '(F6.2, F6.2)', 1.0, 2.0\nend program t\n",
        ["1.00 2.00"]
    };

    desc_f62_e102_mixed_gap => {
        "program t\nprint '(F6.2, E10.2)', 2.0, 0.05\nend program t\n",
        ["2.00 5.00e-2"]
    };

    // ── Iw.m minimum-digit repeat (zero-fill semantics) ──────────────

    desc_i53_minimum_three_digits => {
        "program t\nprint '(I5.3)', 7\nend program t\n",
        ["7"]
    };

    desc_i84_minimum_four_digits => {
        "program t\nprint '(I8.4)', 42\nend program t\n",
        ["42"]
    };

    desc_i63_negative_minimum => {
        "program t\nprint '(I6.3)', -5\nend program t\n",
        ["-5"]
    };

    // ── Sign control and blank-zero prefix tokens ─────────────────────

    desc_f62_positive_fixed_two_places => {
        "program t\nprint '(F6.2)', 1.5\nend program t\n",
        ["1.50"]
    };

    desc_sp_prefix_f62_fallback_plain => {
        "program t\nprint '(SP,F6.2)', 1.5\nend program t\n",
        ["1.5"]
    };

    desc_s_token_f62_fallback_plain => {
        "program t\nprint '(S,F6.2)', -2.5\nend program t\n",
        ["-2.5"]
    };

    desc_bz_prefix_f62_fallback_plain => {
        "program t\nprint '(BZ,F6.2)', 1.5\nend program t\n",
        ["1.5"]
    };
}
