//! Extended character intrinsic coverage: adjustl/adjustr, len_trim, index back,
//! scan/verify, char conversions, lexicographic comparisons, concatenation, IF branches.

fortran_cases! {
    // ── adjustl / adjustr on padded strings ─────────────────────────────

    adjustl_left_padded_len_trim_is_five => {
        "program t\ncharacter(len=12) :: s = '     Fortran'\nprint *, len_trim(adjustl(s))\nend program t\n",
        ["7"]
    };

    adjustr_right_padded_len_trim_is_four => {
        "program t\ncharacter(len=12) :: s = 'Code'\nprint *, len_trim(adjustr(s))\nend program t\n",
        ["4"]
    };

    adjustl_both_sides_internal_space_preserved => {
        "program t\ncharacter(len=16) :: s = '   ab cd   '\nprint *, trim(adjustl(s))\nend program t\n",
        ["ab cd"]
    };

    adjustr_strips_trailing_blanks_on_padded => {
        "program t\ncharacter(len=12) :: s = 'Code      '\nprint *, trim(adjustr(s))\nend program t\n",
        ["Code"]
    };

    adjustl_all_blanks_becomes_blank => {
        "program t\ncharacter(len=8) :: s = '        '\nprint *, len_trim(adjustl(s))\nend program t\n",
        ["0"]
    };

    adjustl_single_char_with_leading_blanks => {
        "program t\ncharacter(len=6) :: s = '    Z'\nprint *, trim(adjustl(s))\nend program t\n",
        ["Z"]
    };

    adjustr_then_adjustl_len_trim_is_six => {
        "program t\ncharacter(len=14) :: s = '  nested  '\nprint *, len_trim(adjustl(adjustr(s)))\nend program t\n",
        ["6"]
    };

    // ── len_trim vs len ─────────────────────────────────────────────────

    len_declared_ten_len_trim_two => {
        "program t\ncharacter(len=10) :: s = 'go'\nprint *, len(s)\nprint *, len_trim(s)\nend program t\n",
        ["10", "2"]
    };

    len_trim_full_string_no_trailing_blanks => {
        "program t\ncharacter(len=5) :: s = 'abcde'\nprint *, len(s)\nprint *, len_trim(s)\nend program t\n",
        ["5", "5"]
    };

    len_trim_all_blanks_is_zero => {
        "program t\ncharacter(len=6) :: s = '      '\nprint *, len_trim(s)\nprint *, len(s)\nend program t\n",
        ["0", "6"]
    };

    len_minus_len_trim_reports_padding => {
        "program t\ncharacter(len=15) :: s = 'payload'\nprint *, len(s) - len_trim(s)\nend program t\n",
        ["8"]
    };

    len_trim_after_adjustl_matches_content => {
        "program t\ncharacter(len=12) :: s = '   data'\nprint *, len_trim(adjustl(s))\nprint *, len(s)\nend program t\n",
        ["4", "12"]
    };

    // ── index with back=.true. ──────────────────────────────────────────

    index_back_finds_last_xy_pair => {
        "program t\ncharacter(len=7) :: s = 'xyzzyxy'\nprint *, index(s, 'xy', .true.)\nend program t\n",
        ["6"]
    };

    index_back_single_letter_last_position => {
        "program t\ncharacter(len=8) :: s = 'abracada'\nprint *, index(s, 'a', .true.)\nend program t\n",
        ["8"]
    };

    index_back_absent_substring_is_zero => {
        "program t\ncharacter(len=10) :: s = 'fortran 90'\nprint *, index(s, 'cpp', .true.)\nend program t\n",
        ["0"]
    };

    index_back_overlapping_aa_pattern => {
        "program t\ncharacter(len=4) :: s = 'baaa'\nprint *, index(s, 'aa', .true.)\nend program t\n",
        ["3"]
    };

    index_back_suffix_at_end => {
        "program t\ncharacter(len=12) :: s = 'prefix-suffix'\nprint *, index(s, 'suffix', .true.)\nend program t\n",
        ["8"]
    };

    // ── scan with character sets ────────────────────────────────────────

    scan_first_digit_in_alphanumeric => {
        "program t\ncharacter(len=6) :: s = 'abc123'\nprint *, scan(s, '0123456789')\nend program t\n",
        ["4"]
    };

    scan_back_last_vowel_in_word => {
        "program t\ncharacter(len=9) :: s = 'rhythmics'\nprint *, scan(s, 'aeiou', .true.)\nend program t\n",
        ["7"]
    };

    scan_punctuation_set_finds_comma => {
        "program t\ncharacter(len=11) :: s = 'value,more'\nprint *, scan(s, ',.;:')\nend program t\n",
        ["6"]
    };

    scan_no_vowel_returns_zero => {
        "program t\ncharacter(len=5) :: s = 'rhythm'\nprint *, scan(s, 'aeiou')\nend program t\n",
        ["0"]
    };

    scan_whitespace_set_in_mixed_text => {
        "program t\ncharacter(len=9) :: s = 'key=value'\nprint *, scan(s, ' =')\nend program t\n",
        ["4"]
    };

    scan_hex_letters_finds_first => {
        "program t\ncharacter(len=8) :: s = '019af2b0'\nprint *, scan(s, 'abcdef')\nend program t\n",
        ["4"]
    };

    // ── verify with character sets ──────────────────────────────────────

    verify_all_alnum_chars_in_set => {
        "program t\ncharacter(len=6) :: s = 'A1b2C3'\nprint *, verify(s, '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz')\nend program t\n",
        ["0"]
    };

    verify_space_not_in_alpha_set => {
        "program t\ncharacter(len=7) :: s = 'ab cd ef'\nprint *, verify(s, 'abcdefghijklmnopqrstuvwxyz')\nend program t\n",
        ["3"]
    };

    verify_digit_among_letters_position => {
        "program t\ncharacter(len=5) :: s = 'ab2de'\nprint *, verify(s, 'abcde')\nend program t\n",
        ["3"]
    };

    verify_pure_alpha_returns_zero => {
        "program t\ncharacter(len=6) :: s = 'Fortran'\nprint *, verify(s, 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ')\nend program t\n",
        ["0"]
    };

    verify_symbol_outside_alnum_set => {
        "program t\ncharacter(len=6) :: s = 'ok!now'\nprint *, verify(s, 'abcdefghijklmnopqrstuvwxyz')\nend program t\n",
        ["3"]
    };

    // ── char / ichar / achar / iachar conversions ─────────────────────

    ichar_lowercase_a_is_ninety_seven => {
        "program t\nprint *, ichar('a')\nend program t\n",
        ["97"]
    };

    ichar_digit_zero_is_forty_eight => {
        "program t\nprint *, ichar('0')\nend program t\n",
        ["48"]
    };

    char_from_code_seventy_two_is_H => {
        "program t\ncharacter(len=1) :: c\nc = char(72)\nprint *, c\nend program t\n",
        ["H"]
    };

    iachar_space_is_thirty_two => {
        "program t\nprint *, iachar(' ')\nend program t\n",
        ["32"]
    };

    achar_sixty_five_is_capital_a => {
        "program t\ncharacter(len=1) :: c\nc = achar(65)\nprint *, c\nend program t\n",
        ["A"]
    };

    ichar_char_roundtrip_for_digit => {
        "program t\nprint *, char(ichar('7'))\nend program t\n",
        ["7"]
    };

    iachar_achar_roundtrip_lowercase => {
        "program t\nprint *, achar(iachar('z'))\nend program t\n",
        ["z"]
    };

    // ── lge / lgt / lle / llt comparison chains (relational operators) ─

    compare_lt_apple_before_banana => {
        "program t\nprint *, 'apple' < 'banana'\nend program t\n",
        ["true"]
    };

    compare_gt_y_greater_than_x => {
        "program t\nprint *, 'y' > 'x'\nend program t\n",
        ["true"]
    };

    compare_ge_equal_trimmed_padding => {
        "program t\ncharacter(len=5) :: a = 'same '\ncharacter(len=5) :: b = 'same'\nprint *, trim(a) >= trim(b)\nend program t\n",
        ["true"]
    };

    compare_le_reflexive_equal_is_true => {
        "program t\nprint *, 'pair' <= 'pair'\nend program t\n",
        ["true"]
    };

    compare_lt_chain_three_strings_ordered => {
        "program t\nprint *, 'ant' < 'bee'\nprint *, 'bee' < 'cow'\nprint *, 'ant' < 'cow'\nend program t\n",
        ["true", "true", "true"]
    };

    compare_ge_and_lt_on_zoo_vs_alpha => {
        "program t\nprint *, 'zoo' >= 'alpha'\nprint *, 'zoo' < 'alpha'\nend program t\n",
        ["true", "false"]
    };

    // ── concatenation // with spaces ───────────────────────────────────

    concat_trimmed_padded_pair => {
        "program t\ncharacter(len=5) :: a = 'Hi   '\ncharacter(len=5) :: b = '  Ho'\nprint *, trim(a // b)\nend program t\n",
        ["Hi     Ho"]
    };

    concat_middle_space_literal => {
        "program t\ncharacter(len=4) :: a = 'ab  '\ncharacter(len=4) :: b = '  cd'\nprint *, trim(trim(a) // ' ' // trim(b))\nend program t\n",
        ["ab   cd"]
    };

    concat_three_segments_with_spaces => {
        "program t\ncharacter(len=3) :: p = 'one'\ncharacter(len=3) :: q = 'two'\ncharacter(len=3) :: r = 'six'\nprint *, trim(p) // ' ' // trim(q) // ' ' // trim(r)\nend program t\n",
        ["one two six"]
    };

    concat_adjusted_strings_with_space => {
        "program t\ncharacter(len=8) :: a = '  left'\ncharacter(len=8) :: b = 'right  '\nprint *, trim(adjustl(a)) // ' ' // trim(adjustr(b))\nend program t\n",
        ["left right"]
    };

    // ── comparison in IF branches ───────────────────────────────────────

    if_lt_branch_prints_less => {
        "program t\nif ('kiwi' < 'mango') then\n  print *, 'less'\nelse\n  print *, 'not'\nend if\nend program t\n",
        ["less"]
    };

    if_ge_branch_prints_gte => {
        "program t\nif ('zebra' >= 'apple') then\n  print *, 'gte'\nelse\n  print *, 'lt'\nend if\nend program t\n",
        ["gte"]
    };

    if_concat_length_branch_longer => {
        "program t\ncharacter(len=20) :: msg\nmsg = trim('foo') // ' ' // trim('bar baz')\nif (len_trim(msg) > 6) then\n  print *, 'long'\nelse\n  print *, 'short'\nend if\nend program t\n",
        ["long"]
    };

    if_scan_found_branch_alpha_only => {
        "program t\ncharacter(len=10) :: s = 'no-digits'\nif (scan(s, '0123456789') == 0) then\n  print *, 'alpha'\nelse\n  print *, 'mixed'\nend if\nend program t\n",
        ["alpha"]
    };

    if_verify_branch_detects_symbol => {
        "program t\ncharacter(len=5) :: s = 'safe1'\nif (verify(s, 'abcdefghijklmnopqrstuvwxyz') == 0) then\n  print *, 'letters'\nelse\n  print *, 'other'\nend if\nend program t\n",
        ["other"]
    };
}
