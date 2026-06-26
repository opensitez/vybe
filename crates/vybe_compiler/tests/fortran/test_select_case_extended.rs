//! Extended SELECT CASE coverage: character selectors, ranges, multiple values,
//! default branch, nested select, computed selectors, and no-fallthrough semantics.

fortran_cases! {
    // ── Integer exact match ──────────────────────────────────────────

    case_int_exact_three => {
        "program t\ninteger :: x = 3\nselect case (x)\ncase (1)\nprint *, \"one\"\ncase (2)\nprint *, \"two\"\ncase (3)\nprint *, \"three\"\ncase default\nprint *, \"other\"\nend select\nend program t\n",
        ["three"]
    };

    case_int_exact_negative => {
        "program t\ninteger :: x = -7\nselect case (x)\ncase (-10:-5)\nprint *, \"negative band\"\ncase (0)\nprint *, \"zero\"\ncase default\nprint *, \"other\"\nend select\nend program t\n",
        ["negative band"]
    };

    case_int_zero_match => {
        "program t\ninteger :: x = 0\nselect case (x)\ncase (0)\nprint *, \"zero\"\ncase (1:100)\nprint *, \"positive\"\ncase default\nprint *, \"other\"\nend select\nend program t\n",
        ["zero"]
    };

    // ── Range selectors ──────────────────────────────────────────────

    case_range_inclusive_low => {
        "program t\ninteger :: n = 1\nselect case (n)\ncase (1:5)\nprint *, \"low\"\ncase (6:10)\nprint *, \"high\"\nend select\nend program t\n",
        ["low"]
    };

    case_range_inclusive_high => {
        "program t\ninteger :: n = 5\nselect case (n)\ncase (1:5)\nprint *, \"low\"\ncase (6:10)\nprint *, \"high\"\nend select\nend program t\n",
        ["low"]
    };

    case_range_between => {
        "program t\ninteger :: n = 8\nselect case (n)\ncase (1:5)\nprint *, \"low\"\ncase (6:10)\nprint *, \"high\"\nend select\nend program t\n",
        ["high"]
    };

    case_range_to_default => {
        "program t\ninteger :: n = 15\nselect case (n)\ncase (1:5)\nprint *, \"low\"\ncase (6:10)\nprint *, \"high\"\ncase default\nprint *, \"out of range\"\nend select\nend program t\n",
        ["out of range"]
    };

    case_open_upper_zero => {
        "program t\ninteger :: n = 0\nselect case (n)\ncase (-10:0)\nprint *, \"non-positive\"\ncase (1:1000)\nprint *, \"positive\"\nend select\nend program t\n",
        ["non-positive"]
    };

    case_open_upper_negative => {
        "program t\ninteger :: n = -3\nselect case (n)\ncase (-10:0)\nprint *, \"non-positive\"\ncase (1:1000)\nprint *, \"positive\"\nend select\nend program t\n",
        ["non-positive"]
    };

    case_open_lower_one => {
        "program t\ninteger :: n = 1\nselect case (n)\ncase (-10:0)\nprint *, \"non-positive\"\ncase (1:1000)\nprint *, \"positive\"\nend select\nend program t\n",
        ["positive"]
    };

    case_open_lower_large => {
        "program t\ninteger :: n = 100\nselect case (n)\ncase (-10:0)\nprint *, \"non-positive\"\ncase (1:1000)\nprint *, \"positive\"\nend select\nend program t\n",
        ["positive"]
    };

    case_single_digit_range => {
        "program t\ninteger :: n = 5\nselect case (n)\ncase (0:9)\nprint *, \"single digit\"\ncase (10:99)\nprint *, \"double digit\"\ncase (100:999)\nprint *, \"triple digit or more\"\nend select\nend program t\n",
        ["single digit"]
    };

    case_double_digit_range => {
        "program t\ninteger :: n = 50\nselect case (n)\ncase (0:9)\nprint *, \"single digit\"\ncase (10:99)\nprint *, \"double digit\"\ncase (100:999)\nprint *, \"triple digit or more\"\nend select\nend program t\n",
        ["double digit"]
    };

    case_large_range_thousands => {
        "program t\ninteger :: n = 5000\nselect case (n)\ncase (1:999)\nprint *, \"hundreds\"\ncase (1000:9999)\nprint *, \"thousands\"\ncase (10000:99999)\nprint *, \"ten-thousands plus\"\nend select\nend program t\n",
        ["thousands"]
    };

    case_boundary_le_four => {
        "program t\ninteger :: n = 4\nselect case (n)\ncase (1:4)\nprint *, \"le 4\"\ncase (5:6)\nprint *, \"ge 5\"\nend select\nend program t\n",
        ["le 4"]
    };

    case_boundary_ge_six => {
        "program t\ninteger :: n = 6\nselect case (n)\ncase (1:4)\nprint *, \"le 4\"\ncase (5:6)\nprint *, \"ge 5\"\nend select\nend program t\n",
        ["ge 5"]
    };

    // ── Multiple values per case ─────────────────────────────────────

    case_multi_value_first => {
        "program t\ninteger :: n = 1\nselect case (n)\ncase (1, 3, 5, 7, 9)\nprint *, \"odd\"\ncase (2, 4, 6, 8, 10)\nprint *, \"even\"\nend select\nend program t\n",
        ["odd"]
    };

    case_multi_value_even_four => {
        "program t\ninteger :: n = 4\nselect case (n)\ncase (1, 3, 5, 7, 9)\nprint *, \"odd\"\ncase (2, 4, 6, 8, 10)\nprint *, \"even\"\nend select\nend program t\n",
        ["even"]
    };

    case_mix_values_range_small => {
        "program t\ninteger :: n = 2\nselect case (n)\ncase (0, 1, 2)\nprint *, \"small\"\ncase (3:10)\nprint *, \"medium\"\ncase (11:99)\nprint *, \"large\"\nend select\nend program t\n",
        ["small"]
    };

    case_mix_values_range_medium => {
        "program t\ninteger :: n = 7\nselect case (n)\ncase (0, 1, 2)\nprint *, \"small\"\ncase (3:10)\nprint *, \"medium\"\ncase (11:99)\nprint *, \"large\"\nend select\nend program t\n",
        ["medium"]
    };

    case_mix_values_range_large => {
        "program t\ninteger :: n = 20\nselect case (n)\ncase (0, 1, 2)\nprint *, \"small\"\ncase (3:10)\nprint *, \"medium\"\ncase (11:99)\nprint *, \"large\"\nend select\nend program t\n",
        ["large"]
    };

    case_multi_values_loop_match => {
        "program t\ninteger :: i\ndo i = 1, 6\nselect case (i)\ncase (1, 2, 6)\nprint *, \"match\"\ncase default\nprint *, \"no\"\nend select\nend do\nend program t\n",
        ["match", "match", "no", "no", "no", "match"]
    };

    // ── Default branch ───────────────────────────────────────────────

    case_default_unmatched => {
        "program t\ninteger :: n = -1\nselect case (n)\ncase (1:10)\nprint *, \"in range\"\ncase default\nprint *, \"fallback\"\nend select\nend program t\n",
        ["fallback"]
    };

    case_default_only => {
        "program t\ninteger :: n = 42\nselect case (n)\ncase default\nprint *, \"default\"\nend select\nend program t\n",
        ["default"]
    };

    case_default_after_ranges_miss => {
        "program t\ninteger :: n = 0\nselect case (n)\ncase (1, 2)\nprint *, \"listed\"\ncase default\nprint *, \"other\"\nend select\nend program t\n",
        ["other"]
    };

    // ── Character selectors ──────────────────────────────────────────

    case_char_exact_b => {
        "program t\ncharacter :: c = 'b'\nselect case (c)\ncase ('a')\nprint *, \"a\"\ncase ('b')\nprint *, \"b\"\ncase ('c')\nprint *, \"c\"\nend select\nend program t\n",
        ["b"]
    };

    case_char_range_first_half => {
        "program t\ncharacter :: c = 'f'\nselect case (c)\ncase ('a':'m')\nprint *, \"first half\"\ncase ('n':'z')\nprint *, \"second half\"\nend select\nend program t\n",
        ["first half"]
    };

    case_char_range_second_half => {
        "program t\ncharacter :: c = 't'\nselect case (c)\ncase ('a':'m')\nprint *, \"first half\"\ncase ('n':'z')\nprint *, \"second half\"\nend select\nend program t\n",
        ["second half"]
    };

    case_char_uppercase_z => {
        "program t\ncharacter :: c = 'Z'\nselect case (c)\ncase ('A':'Z')\nprint *, \"uppercase\"\ncase ('a':'z')\nprint *, \"lowercase\"\ncase default\nprint *, \"other\"\nend select\nend program t\n",
        ["uppercase"]
    };

    case_char_vowel_e => {
        "program t\ncharacter :: c = 'e'\nselect case (c)\ncase ('a', 'e', 'i', 'o', 'u')\nprint *, \"vowel\"\ncase default\nprint *, \"consonant\"\nend select\nend program t\n",
        ["vowel"]
    };

    case_char_consonant_b => {
        "program t\ncharacter :: c = 'b'\nselect case (c)\ncase ('a', 'e', 'i', 'o', 'u')\nprint *, \"vowel\"\ncase default\nprint *, \"consonant\"\nend select\nend program t\n",
        ["consonant"]
    };

    case_char_string_foo => {
        "program t\ncharacter(len=3) :: s = 'foo'\nselect case (s)\ncase ('bar')\nprint *, \"bar\"\ncase ('baz')\nprint *, \"baz\"\ncase ('foo')\nprint *, \"foo\"\ncase default\nprint *, \"other\"\nend select\nend program t\n",
        ["foo"]
    };

    case_char_string_other => {
        "program t\ncharacter(len=3) :: s = 'xyz'\nselect case (s)\ncase ('bar')\nprint *, \"bar\"\ncase ('baz')\nprint *, \"baz\"\ncase ('foo')\nprint *, \"foo\"\ncase default\nprint *, \"other\"\nend select\nend program t\n",
        ["other"]
    };

    // ── Nested SELECT CASE ───────────────────────────────────────────

    nested_outer_i_equals_one => {
        "program t\ninteger :: i = 1, j = 3\nselect case (i)\ncase (1)\nprint *, \"outer one\"\ncase (2)\nselect case (j)\ncase (1:2)\nprint *, \"inner small\"\ncase (3:10)\nprint *, \"inner large\"\nend select\ncase default\nprint *, \"outer other\"\nend select\nend program t\n",
        ["outer one"]
    };

    nested_inner_small => {
        "program t\ninteger :: i = 2, j = 1\nselect case (i)\ncase (1)\nprint *, \"outer one\"\ncase (2)\nselect case (j)\ncase (1:2)\nprint *, \"inner small\"\ncase (3:10)\nprint *, \"inner large\"\nend select\ncase default\nprint *, \"outer other\"\nend select\nend program t\n",
        ["inner small"]
    };

    nested_inner_large => {
        "program t\ninteger :: i = 2, j = 5\nselect case (i)\ncase (1)\nprint *, \"outer one\"\ncase (2)\nselect case (j)\ncase (1:2)\nprint *, \"inner small\"\ncase (3:10)\nprint *, \"inner large\"\nend select\ncase default\nprint *, \"outer other\"\nend select\nend program t\n",
        ["inner large"]
    };

    nested_outer_default => {
        "program t\ninteger :: i = 9, j = 1\nselect case (i)\ncase (1)\nprint *, \"outer one\"\ncase (2)\nselect case (j)\ncase (1:2)\nprint *, \"inner small\"\ncase (3:10)\nprint *, \"inner large\"\nend select\ncase default\nprint *, \"outer other\"\nend select\nend program t\n",
        ["outer other"]
    };

    nested_select_in_loop => {
        "program t\ninteger :: i, j\ndo i = 1, 3\nselect case (i)\ncase (1)\ndo j = 1, 2\nselect case (j)\ncase (1)\nprint *, \"one-one\"\ncase (2)\nprint *, \"one-two\"\nend select\nend do\ncase (2:3)\nprint *, \"outer range\"\nend select\nend do\nend program t\n",
        ["one-one", "one-two", "outer range", "outer range"]
    };

    // ── Computed selector expressions ────────────────────────────────

    case_expr_small_sum => {
        "program t\ninteger :: x = 2, y = 3\nselect case (x + y)\ncase (1:7)\nprint *, \"small sum\"\ncase (8:10)\nprint *, \"medium sum\"\ncase (11:20)\nprint *, \"large sum\"\nend select\nend program t\n",
        ["small sum"]
    };

    case_expr_medium_sum => {
        "program t\ninteger :: x = 5, y = 4\nselect case (x + y)\ncase (1:7)\nprint *, \"small sum\"\ncase (8:10)\nprint *, \"medium sum\"\ncase (11:20)\nprint *, \"large sum\"\nend select\nend program t\n",
        ["medium sum"]
    };

    case_expr_large_sum => {
        "program t\ninteger :: x = 10, y = 5\nselect case (x + y)\ncase (1:7)\nprint *, \"small sum\"\ncase (8:10)\nprint *, \"medium sum\"\ncase (11:20)\nprint *, \"large sum\"\nend select\nend program t\n",
        ["large sum"]
    };

    case_on_sum_array => {
        "program t\ninteger :: a(5) = [1, 2, 3, 4, 5]\nselect case (sum(a))\ncase (1:10)\nprint *, \"small\"\ncase (11:20)\nprint *, \"medium\"\ncase (21:30)\nprint *, \"large\"\nend select\nend program t\n",
        ["medium"]
    };

    case_on_mod_zero => {
        "program t\ninteger :: i = 6\nselect case (mod(i, 3))\ncase (0)\nprint *, \"rem zero\"\ncase (1)\nprint *, \"rem one\"\ncase (2)\nprint *, \"rem two\"\nend select\nend program t\n",
        ["rem zero"]
    };

    case_on_mod_one => {
        "program t\ninteger :: i = 7\nselect case (mod(i, 3))\ncase (0)\nprint *, \"rem zero\"\ncase (1)\nprint *, \"rem one\"\ncase (2)\nprint *, \"rem two\"\nend select\nend program t\n",
        ["rem one"]
    };

    case_on_abs_negative => {
        "program t\ninteger :: n = -9\nselect case (abs(n))\ncase (1:5)\nprint *, \"small abs\"\ncase (6:10)\nprint *, \"large abs\"\nend select\nend program t\n",
        ["large abs"]
    };

    case_on_product => {
        "program t\ninteger :: a = 3, b = 4\nselect case (a * b)\ncase (1:10)\nprint *, \"small product\"\ncase (11:20)\nprint *, \"large product\"\nend select\nend program t\n",
        ["small product"]
    };

    // ── No fall-through (each case independent) ────────────────────

    case_no_fallthrough_first => {
        "program t\ninteger :: n = 1\nselect case (n)\ncase (1)\nprint *, \"first\"\ncase (2)\nprint *, \"second\"\ncase (3)\nprint *, \"third\"\nend select\nprint *, \"after\"\nend program t\n",
        ["first", "after"]
    };

    case_no_fallthrough_second => {
        "program t\ninteger :: n = 2\nselect case (n)\ncase (1)\nprint *, \"first\"\ncase (2)\nprint *, \"second\"\ncase (3)\nprint *, \"third\"\nend select\nprint *, \"after\"\nend program t\n",
        ["second", "after"]
    };

    case_no_fallthrough_third => {
        "program t\ninteger :: n = 3\nselect case (n)\ncase (1)\nprint *, \"first\"\ncase (2)\nprint *, \"second\"\ncase (3)\nprint *, \"third\"\nend select\nprint *, \"after\"\nend program t\n",
        ["third", "after"]
    };

    case_no_match_no_default_skips => {
        "program t\ninteger :: n = 99\nselect case (n)\ncase (1)\nprint *, \"one\"\ncase (2)\nprint *, \"two\"\nend select\nprint *, \"done\"\nend program t\n",
        ["done"]
    };

    case_independent_per_iteration => {
        "program t\ninteger :: i\ndo i = 1, 4\nselect case (i)\ncase (1)\nprint *, \"alpha\"\ncase (2)\nprint *, \"beta\"\ncase (3)\nprint *, \"gamma\"\ncase default\nprint *, \"omega\"\nend select\nend do\nend program t\n",
        ["alpha", "beta", "gamma", "omega"]
    };
}
