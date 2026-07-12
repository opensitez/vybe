//! Extended lexical comparisons: llt, lgt, lle, lge with case and padding patterns.
//! Distinct from relational operators in `test_character_intrinsics_extended.rs`.

fortran_cases! {
    llt_alpha_before_beta => {
        "program t\nprint *, llt('alpha', 'beta')\nend program t\n",
        ["true"]
    };

    llt_equal_strings_false => {
        "program t\nprint *, llt('same', 'same')\nend program t\n",
        ["false"]
    };

    llt_shorter_prefix_before_longer => {
        "program t\nprint *, llt('abc', 'abcd')\nend program t\n",
        ["true"]
    };

    llt_uppercase_before_lowercase_ascii => {
        "program t\nprint *, llt('A', 'a')\nend program t\n",
        ["true"]
    };

    llt_digit_before_letter => {
        "program t\nprint *, llt('1', 'a')\nend program t\n",
        ["true"]
    };

    lgt_beta_after_alpha => {
        "program t\nprint *, lgt('beta', 'alpha')\nend program t\n",
        ["true"]
    };

    lgt_equal_strings_false => {
        "program t\nprint *, lgt('same', 'same')\nend program t\n",
        ["false"]
    };

    lgt_longer_after_prefix => {
        "program t\nprint *, lgt('abcd', 'abc')\nend program t\n",
        ["true"]
    };

    lgt_lowercase_after_uppercase => {
        "program t\nprint *, lgt('a', 'A')\nend program t\n",
        ["true"]
    };

    lgt_letter_after_digit => {
        "program t\nprint *, lgt('a', '1')\nend program t\n",
        ["true"]
    };

    lle_equal_strings_true => {
        "program t\nprint *, lle('pair', 'pair')\nend program t\n",
        ["true"]
    };

    lle_alpha_less_or_equal_beta => {
        "program t\nprint *, lle('alpha', 'beta')\nend program t\n",
        ["true"]
    };

    lle_not_greater_than_equal => {
        "program t\nprint *, lle('beta', 'alpha')\nend program t\n",
        ["false"]
    };

    lle_prefix_less_or_equal => {
        "program t\nprint *, lle('ab', 'abc')\nend program t\n",
        ["true"]
    };

    lle_space_before_letter => {
        "program t\nprint *, lle(' ', 'a')\nend program t\n",
        ["true"]
    };

    lge_equal_strings_true => {
        "program t\nprint *, lge('pair', 'pair')\nend program t\n",
        ["true"]
    };

    lge_beta_greater_or_equal_alpha => {
        "program t\nprint *, lge('beta', 'alpha')\nend program t\n",
        ["true"]
    };

    lge_not_less_than_equal => {
        "program t\nprint *, lge('alpha', 'beta')\nend program t\n",
        ["false"]
    };

    lge_longer_greater_or_equal_prefix => {
        "program t\nprint *, lge('abc', 'ab')\nend program t\n",
        ["true"]
    };

    lge_letter_greater_or_equal_space => {
        "program t\nprint *, lge('a', ' ')\nend program t\n",
        ["true"]
    };

    lex_chain_llt_and_lgt_opposite => {
        "program t\nprint *, llt('a','b') .and. lgt('b','a')\nend program t\n",
        ["true"]
    };

    lex_chain_lle_lge_equal_reflexive => {
        "program t\nprint *, lle('x','x') .and. lge('x','x')\nend program t\n",
        ["true"]
    };

    lex_sort_three_words_count => {
        "program t\ncharacter(len=5) :: w(3) = ['apple','grape','banana']\ninteger :: i, j, c\nc = 0\ndo i = 1, 2\n  do j = i+1, 3\n    if (llt(w(i), w(j))) c = c + 1\n  end do\nend do\nprint *, c\nend program t\n",
        ["3"]
    };

    lex_case_pair_upper_less_than_lower => {
        "program t\nprint *, llt('CAT', 'cat')\nend program t\n",
        ["true"]
    };

    lex_case_pair_lower_greater_upper => {
        "program t\nprint *, lgt('dog', 'DOG')\nend program t\n",
        ["true"]
    };

    lex_digit_string_less_than_alpha => {
        "program t\nprint *, llt('999', 'aaa')\nend program t\n",
        ["true"]
    };

    lex_blank_padded_compare_trimmed => {
        "program t\ncharacter(len=5) :: a = 'hi   '\ncharacter(len=5) :: b = 'hi'\nprint *, lge(a, b)\nend program t\n",
        ["true"]
    };

    lex_select_case_on_lgt => {
        "program t\nselect case (lgt('z', 'a'))\ncase (.true.)\nprint *, 'gt'\ncase default\nprint *, 'no'\nend select\nend program t\n",
        ["gt"]
    };

    lex_if_branch_lle_detects_equal => {
        "program t\nif (lle('ok', 'ok')) then\nprint *, 'eq'\nelse\nprint *, 'ne'\nend if\nend program t\n",
        ["eq"]
    };

    lex_if_branch_llt_detects_less => {
        "program t\nif (llt('ant', 'bee')) then\nprint *, 'less'\nelse\nprint *, 'not'\nend if\nend program t\n",
        ["less"]
    };

    lex_array_max_via_lgt => {
        "program t\ncharacter(len=3) :: a(3) = ['cat','dog','bat']\ncharacter(len=3) :: m\ninteger :: i\nm = a(1)\ndo i = 2, 3\nif (lgt(a(i), m)) m = a(i)\nend do\nprint *, trim(m)\nend program t\n",
        ["dog"]
    };

    lex_array_min_via_llt => {
        "program t\ncharacter(len=3) :: a(3) = ['cat','dog','bat']\ncharacter(len=3) :: m\ninteger :: i\nm = a(1)\ndo i = 2, 3\nif (llt(a(i), m)) m = a(i)\nend do\nprint *, trim(m)\nend program t\n",
        ["bat"]
    };

    lex_neqv_of_comparisons => {
        "program t\nprint *, llt('a','b') .neqv. lgt('a','b')\nend program t\n",
        ["true"]
    };

    lex_eqv_reflexive_lge => {
        "program t\nprint *, lge('test','test') .eqv. lle('test','test')\nend program t\n",
        ["true"]
    };

    lex_mixed_case_word_order => {
        "program t\nprint *, llt('Fortran', 'fortran')\nend program t\n",
        ["true"]
    };

    lex_underscore_vs_letter => {
        "program t\nprint *, llt('a_b', 'ab')\nend program t\n",
        ["true"]
    };

    lex_number_prefix_vs_full => {
        "program t\nprint *, llt('12', '123')\nend program t\n",
        ["true"]
    };

    lex_sign_chars_order => {
        "program t\nprint *, llt('+', '-')\nend program t\n",
        ["true"]
    };

    lex_space_less_than_digit => {
        "program t\nprint *, llt(' ', '0')\nend program t\n",
        ["true"]
    };

    lex_tab_vs_space => {
        "program t\ncharacter(len=1) :: t = char(9)\nprint *, lgt(t, ' ')\nend program t\n",
        ["true"]
    };

    lex_compare_in_merge => {
        "program t\nprint *, merge(1, 0, llt('a', 'b'))\nend program t\n",
        ["1"]
    };

    lex_compare_in_merge_false => {
        "program t\nprint *, merge(1, 0, lgt('a', 'b'))\nend program t\n",
        ["0"]
    };

    lex_four_way_consistency => {
        "program t\nprint *, llt('p','q')\nprint *, lgt('q','p')\nprint *, lle('p','q')\nprint *, lge('q','p')\nend program t\n",
        ["true", "true", "true", "true"]
    };

    lex_equal_neither_lt_nor_gt => {
        "program t\nprint *, llt('eq','eq')\nprint *, lgt('eq','eq')\nend program t\n",
        ["false", "false"]
    };

    lex_transitivity_chain => {
        "program t\nprint *, llt('a','b') .and. llt('b','c') .and. llt('a','c')\nend program t\n",
        ["true"]
    };

    lex_reverse_transitivity_on_gt => {
        "program t\nprint *, lgt('c','b') .and. lgt('b','a') .and. lgt('c','a')\nend program t\n",
        ["true"]
    };

    lex_empty_vs_blank_char => {
        "program t\ncharacter(len=1) :: a = ' '\nprint *, lle(a, ' ')\nend program t\n",
        ["true"]
    };

    lex_punctuation_order => {
        "program t\nprint *, llt('.', ',')\nend program t\n",
        ["true"]
    };

    lex_hex_letter_case => {
        "program t\nprint *, llt('a', 'A')\nend program t\n",
        ["true"]
    };

    lex_year_strings => {
        "program t\nprint *, llt('1999', '2000')\nend program t\n",
        ["true"]
    };

    lex_version_strings => {
        "program t\nprint *, llt('1.09', '1.10')\nend program t\n",
        ["true"]
    };

    lex_country_codes => {
        "program t\nprint *, llt('US', 'USA')\nend program t\n",
        ["true"]
    };

    lex_suffix_sort => {
        "program t\nprint *, llt('file.txt', 'file.txz')\nend program t\n",
        ["true"]
    };

    lex_prefix_shared => {
        "program t\nprint *, lge('prefix', 'pre')\nend program t\n",
        ["true"]
    };

    lex_case_insensitive_simulation => {
        "program t\ncharacter(len=3) :: a = 'AbC'\ncharacter(len=3) :: b = 'aBc'\nprint *, llt(a, b)\nend program t\n",
        ["true"]
    };

}
