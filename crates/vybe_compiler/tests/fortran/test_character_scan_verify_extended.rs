//! Extended scan, verify, index, len_trim, repeat, adjustl/adjustr, achar/ichar.
//! Distinct from `test_character_intrinsics_extended.rs` and `test_strings_extended.rs`.

fortran_cases! {
    scan_forward_first_vowel_in_sentence => {
        "program t\ncharacter(len=15) :: s = 'The quick brown'\nprint *, scan(s, 'aeiou')\nend program t\n",
        ["3"]
    };

    scan_forward_first_digit_in_mixed => {
        "program t\ncharacter(len=9) :: s = 'abc123def'\nprint *, scan(s, '0123456789')\nend program t\n",
        ["4"]
    };

    scan_forward_first_space_in_tokens => {
        "program t\ncharacter(len=7) :: s = 'one two'\nprint *, scan(s, ' ')\nend program t\n",
        ["4"]
    };

    scan_forward_first_upper_in_mixed => {
        "program t\ncharacter(len=6) :: s = 'aBcDeF'\nprint *, scan(s, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ')\nend program t\n",
        ["2"]
    };

    scan_forward_punctuation_comma => {
        "program t\ncharacter(len=10) :: s = 'value,more'\nprint *, scan(s, ',.;:')\nend program t\n",
        ["6"]
    };

    scan_forward_no_match_returns_zero => {
        "program t\ncharacter(len=6) :: s = 'rhythm'\nprint *, scan(s, 'aeiou')\nend program t\n",
        ["0"]
    };

    scan_forward_first_hex_in_uuid => {
        "program t\ncharacter(len=8) :: s = '019af2b0'\nprint *, scan(s, 'abcdef')\nend program t\n",
        ["3"]
    };

    scan_forward_equals_in_pair => {
        "program t\ncharacter(len=9) :: s = 'key=value'\nprint *, scan(s, ' =')\nend program t\n",
        ["4"]
    };

    scan_backward_last_vowel_in_hello => {
        "program t\ncharacter(len=5) :: s = 'hello'\nprint *, scan(s, 'aeiou', .true.)\nend program t\n",
        ["5"]
    };

    scan_backward_last_digit_in_trailing => {
        "program t\ncharacter(len=7) :: s = 'test007'\nprint *, scan(s, '0123456789', .true.)\nend program t\n",
        ["7"]
    };

    scan_backward_last_space_in_padded => {
        "program t\ncharacter(len=7) :: s = 'a b c  '\nprint *, scan(s, ' ', .true.)\nend program t\n",
        ["5"]
    };

    scan_backward_no_vowel_in_consonants => {
        "program t\ncharacter(len=5) :: s = 'bcdfg'\nprint *, scan(s, 'aeiou', .true.)\nend program t\n",
        ["0"]
    };

    scan_backward_last_comma_in_list => {
        "program t\ncharacter(len=6) :: s = 'a,b,c,'\nprint *, scan(s, ',;', .true.)\nend program t\n",
        ["5"]
    };

    scan_backward_last_lower_in_caps => {
        "program t\ncharacter(len=6) :: s = 'ABCdEF'\nprint *, scan(s, 'abcdefghijklmnopqrstuvwxyz', .true.)\nend program t\n",
        ["4"]
    };

    verify_all_alpha_is_zero => {
        "program t\ncharacter(len=8) :: s = 'alphabet'\nprint *, verify(s, 'abcdefghijklmnopqrstuvwxyz')\nend program t\n",
        ["0"]
    };

    verify_first_nonalpha_at_start => {
        "program t\ncharacter(len=4) :: s = '1abc'\nprint *, verify(s, '0123456789')\nend program t\n",
        ["1"]
    };

    verify_space_is_nonalpha => {
        "program t\ncharacter(len=5) :: s = 'ab cd'\nprint *, verify(s, 'abcdefghijklmnopqrstuvwxyz')\nend program t\n",
        ["3"]
    };

    verify_digit_among_letters => {
        "program t\ncharacter(len=5) :: s = 'ab2de'\nprint *, verify(s, 'abcde')\nend program t\n",
        ["3"]
    };

    verify_all_alnum_set => {
        "program t\ncharacter(len=9) :: s = 'Fortran90'\nprint *, verify(s, 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789')\nend program t\n",
        ["0"]
    };

    verify_symbol_in_word => {
        "program t\ncharacter(len=6) :: s = 'ok!now'\nprint *, verify(s, 'abcdefghijklmnopqrstuvwxyz')\nend program t\n",
        ["3"]
    };

    verify_leading_blank_not_in_set => {
        "program t\ncharacter(len=6) :: s = '  data'\nprint *, verify(s, 'abcdefghijklmnopqrstuvwxyz')\nend program t\n",
        ["1"]
    };

    verify_tab_not_in_letters => {
        "program t\ncharacter(len=3) :: s = 'a\tb'\nprint *, verify(s, 'ab')\nend program t\n",
        ["2"]
    };

    index_forward_substring_at_start => {
        "program t\ncharacter(len=7) :: s = 'fortran'\nprint *, index(s, 'for')\nend program t\n",
        ["1"]
    };

    index_forward_substring_middle => {
        "program t\ncharacter(len=7) :: s = 'fortran'\nprint *, index(s, 'tra')\nend program t\n",
        ["3"]
    };

    index_forward_not_found => {
        "program t\ncharacter(len=7) :: s = 'fortran'\nprint *, index(s, 'java')\nend program t\n",
        ["0"]
    };

    index_forward_single_char => {
        "program t\ncharacter(len=11) :: s = 'mississippi'\nprint *, index(s, 's')\nend program t\n",
        ["2"]
    };

    index_forward_overlap_pattern => {
        "program t\ncharacter(len=4) :: s = 'aaaa'\nprint *, index(s, 'aa')\nend program t\n",
        ["1"]
    };

    index_backward_last_bc => {
        "program t\ncharacter(len=6) :: s = 'abcabc'\nprint *, index(s, 'bc')\nend program t\n",
        ["5"]
    };

    index_backward_last_a => {
        "program t\ncharacter(len=8) :: s = 'abracada'\nprint *, index(s, 'a')\nend program t\n",
        ["8"]
    };

    index_backward_not_found => {
        "program t\ncharacter(len=7) :: s = 'fortran'\nprint *, index(s, 'cpp')\nend program t\n",
        ["0"]
    };

    index_backward_suffix => {
        "program t\ncharacter(len=13) :: s = 'prefix-suffix'\nprint *, index(s, 'suffix')\nend program t\n",
        ["8"]
    };

    index_forward_after_position_slice => {
        "program t\ncharacter(len=11) :: s = 'one two one'\nprint *, index(s(5:), 'one')\nend program t\n",
        ["5"]
    };

    len_trim_short_in_long_buffer => {
        "program t\ncharacter(len=10) :: s = 'go        '\nprint *, len_trim(s)\nend program t\n",
        ["2"]
    };

    len_trim_full_no_trailing => {
        "program t\ncharacter(len=5) :: s = 'abcde'\nprint *, len_trim(s)\nend program t\n",
        ["5"]
    };

    len_trim_all_blanks_zero => {
        "program t\ncharacter(len=6) :: s = '      '\nprint *, len_trim(s)\nend program t\n",
        ["0"]
    };

    len_trim_single_char_padded => {
        "program t\ncharacter(len=8) :: s = 'x       '\nprint *, len_trim(s)\nend program t\n",
        ["1"]
    };

    len_trim_internal_spaces_counted => {
        "program t\ncharacter(len=7) :: s = 'a b c  '\nprint *, len_trim(s)\nend program t\n",
        ["5"]
    };

    repeat_dash_three_times => {
        "program t\nprint *, repeat('ab', 3)\nend program t\n",
        ["ababab"]
    };

    repeat_single_char_five => {
        "program t\nprint *, repeat('x', 5)\nend program t\n",
        ["xxxxx"]
    };

    repeat_once_is_identity => {
        "program t\nprint *, repeat('ok', 1)\nend program t\n",
        ["ok"]
    };

    repeat_zero_is_empty => {
        "program t\nprint *, repeat('hi', 0)\nend program t\n",
        [""]
    };

    repeat_star_pattern => {
        "program t\nprint *, repeat('*', 4)\nend program t\n",
        ["****"]
    };

    adjustl_moves_leading_blanks => {
        "program t\ncharacter(len=10) :: s = '   data   '\nprint *, trim(adjustl(s))\nend program t\n",
        ["data"]
    };

    adjustr_moves_trailing_content => {
        "program t\ncharacter(len=8) :: s = 'code    '\nprint *, trim(adjustr(s))\nend program t\n",
        ["code"]
    };

    adjustl_preserves_internal_space => {
        "program t\ncharacter(len=10) :: s = '  ab cd   '\nprint *, trim(adjustl(s))\nend program t\n",
        ["ab cd"]
    };

    adjustr_right_aligns_in_field => {
        "program t\ncharacter(len=6) :: s = 'xy    '\nprint *, trim(adjustr(s))\nend program t\n",
        ["xy"]
    };

    adjustl_all_blanks_stays_blank => {
        "program t\ncharacter(len=5) :: s = '     '\nprint *, trim(adjustl(s))\nend program t\n",
        ["     "]
    };

    adjustl_then_len_trim => {
        "program t\ncharacter(len=6) :: s = '   z  '\nprint *, len_trim(adjustl(s))\nend program t\n",
        ["1"]
    };

    ichar_digit_0 => {
        "program t\nprint *, ichar('0')\nend program t\n",
        ["48"]
    };

    ichar_digit_1 => {
        "program t\nprint *, ichar('1')\nend program t\n",
        ["49"]
    };

    ichar_digit_2 => {
        "program t\nprint *, ichar('2')\nend program t\n",
        ["50"]
    };

    ichar_digit_3 => {
        "program t\nprint *, ichar('3')\nend program t\n",
        ["51"]
    };

    ichar_digit_4 => {
        "program t\nprint *, ichar('4')\nend program t\n",
        ["52"]
    };

    ichar_digit_5 => {
        "program t\nprint *, ichar('5')\nend program t\n",
        ["53"]
    };

    ichar_digit_6 => {
        "program t\nprint *, ichar('6')\nend program t\n",
        ["54"]
    };

    ichar_digit_7 => {
        "program t\nprint *, ichar('7')\nend program t\n",
        ["55"]
    };

    ichar_digit_8 => {
        "program t\nprint *, ichar('8')\nend program t\n",
        ["56"]
    };

    ichar_digit_9 => {
        "program t\nprint *, ichar('9')\nend program t\n",
        ["57"]
    };

    ichar_letter_a => {
        "program t\nprint *, ichar('A')\nend program t\n",
        ["65"]
    };

    ichar_letter_m => {
        "program t\nprint *, ichar('M')\nend program t\n",
        ["77"]
    };

    ichar_letter_z => {
        "program t\nprint *, ichar('Z')\nend program t\n",
        ["90"]
    };

    ichar_letter_a => {
        "program t\nprint *, ichar('a')\nend program t\n",
        ["97"]
    };

    ichar_letter_m => {
        "program t\nprint *, ichar('m')\nend program t\n",
        ["109"]
    };

    ichar_letter_z => {
        "program t\nprint *, ichar('z')\nend program t\n",
        ["122"]
    };

    achar_code_32 => {
        "program t\ncharacter(len=1) :: c\nc = achar(32)\nprint *, ichar(c)\nend program t\n",
        ["32"]
    };

    achar_code_33 => {
        "program t\ncharacter(len=1) :: c\nc = achar(33)\nprint *, ichar(c)\nend program t\n",
        ["33"]
    };

    achar_code_64 => {
        "program t\ncharacter(len=1) :: c\nc = achar(64)\nprint *, ichar(c)\nend program t\n",
        ["64"]
    };

}
