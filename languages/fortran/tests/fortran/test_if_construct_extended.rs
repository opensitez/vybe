//! Extended IF construct coverage: multi-branch ELSEIF chains, block IF without
//! ELSE, compound logical conditions, character comparisons, and legacy arithmetic IF.

fortran_cases! {
    // ── IF-THEN-ELSEIF chains ─────────────────────────────────────────

    if_elif_chain_hits_first_of_four => {
        "program t\ninteger :: code = 10\nif (code == 10) then\nprint *, \"ten\"\nelse if (code == 20) then\nprint *, \"twenty\"\nelse if (code == 30) then\nprint *, \"thirty\"\nelse\nprint *, \"other\"\nend if\nend program t\n",
        ["ten"]
    };

    if_elif_chain_hits_second_of_four => {
        "program t\ninteger :: code = 20\nif (code == 10) then\nprint *, \"ten\"\nelse if (code == 20) then\nprint *, \"twenty\"\nelse if (code == 30) then\nprint *, \"thirty\"\nelse\nprint *, \"other\"\nend if\nend program t\n",
        ["twenty"]
    };

    if_elif_chain_hits_third_of_four => {
        "program t\ninteger :: code = 30\nif (code == 10) then\nprint *, \"ten\"\nelse if (code == 20) then\nprint *, \"twenty\"\nelse if (code == 30) then\nprint *, \"thirty\"\nelse\nprint *, \"other\"\nend if\nend program t\n",
        ["thirty"]
    };

    if_elif_chain_falls_through_to_else => {
        "program t\ninteger :: code = 99\nif (code == 10) then\nprint *, \"ten\"\nelse if (code == 20) then\nprint *, \"twenty\"\nelse if (code == 30) then\nprint *, \"thirty\"\nelse\nprint *, \"other\"\nend if\nend program t\n",
        ["other"]
    };

    if_elif_chain_keeps_first_true => {
        "program t\ninteger :: code = 2\nif (code == 1) then\nprint *, \"one\"\nelse if (code == 2) then\nprint *, \"two\"\nelse if (code == 2 .or. code == 3) then\nprint *, \"also-two\"\nelse\nprint *, \"other\"\nend if\nend program t\n",
        ["two"]
    };

    if_elif_temperature_freezing_branch => {
        "program t\nreal :: t = -5.0\nif (t < 0.0) then\nprint *, \"freezing\"\nelse if (t < 15.0) then\nprint *, \"cool\"\nelse if (t < 25.0) then\nprint *, \"mild\"\nelse\nprint *, \"warm\"\nend if\nend program t\n",
        ["freezing"]
    };

    if_elif_temperature_mild_branch => {
        "program t\nreal :: t = 20.0\nif (t < 0.0) then\nprint *, \"freezing\"\nelse if (t < 15.0) then\nprint *, \"cool\"\nelse if (t < 25.0) then\nprint *, \"mild\"\nelse\nprint *, \"warm\"\nend if\nend program t\n",
        ["mild"]
    };

    if_elif_mod_three_prints_fizz => {
        "program t\ninteger :: n = 9\nif (mod(n, 15) == 0) then\nprint *, \"fizzbuzz\"\nelse if (mod(n, 3) == 0) then\nprint *, \"fizz\"\nelse if (mod(n, 5) == 0) then\nprint *, \"buzz\"\nelse\nprint *, n\nend if\nend program t\n",
        ["fizz"]
    };

    if_elif_mod_five_prints_buzz => {
        "program t\ninteger :: n = 10\nif (mod(n, 15) == 0) then\nprint *, \"fizzbuzz\"\nelse if (mod(n, 3) == 0) then\nprint *, \"fizz\"\nelse if (mod(n, 5) == 0) then\nprint *, \"buzz\"\nelse\nprint *, n\nend if\nend program t\n",
        ["buzz"]
    };

    if_elif_weekday_selects_wednesday => {
        "program t\ninteger :: day = 3\nif (day == 1) then\nprint *, \"mon\"\nelse if (day == 2) then\nprint *, \"tue\"\nelse if (day == 3) then\nprint *, \"wed\"\nelse if (day == 4) then\nprint *, \"thu\"\nelse if (day == 5) then\nprint *, \"fri\"\nelse\nprint *, \"weekend\"\nend if\nend program t\n",
        ["wed"]
    };

    if_elif_sign_negative_branch => {
        "program t\ninteger :: v = -12\nif (v > 0) then\nprint *, \"pos\"\nelse if (v < 0) then\nprint *, \"neg\"\nelse\nprint *, \"zero\"\nend if\nend program t\n",
        ["neg"]
    };

    // ── Block IF with multiple ELSEIF arms ──────────────────────────────

    block_if_five_tier_grade_a_plus => {
        "program t\ninteger :: pts = 98\nif (pts >= 97) then\nprint *, \"A+\"\nelse if (pts >= 93) then\nprint *, \"A\"\nelse if (pts >= 90) then\nprint *, \"A-\"\nelse if (pts >= 87) then\nprint *, \"B+\"\nelse\nprint *, \"B-\"\nend if\nend program t\n",
        ["A+"]
    };

    block_if_five_tier_grade_a_minus => {
        "program t\ninteger :: pts = 91\nif (pts >= 97) then\nprint *, \"A+\"\nelse if (pts >= 93) then\nprint *, \"A\"\nelse if (pts >= 90) then\nprint *, \"A-\"\nelse if (pts >= 87) then\nprint *, \"B+\"\nelse\nprint *, \"B-\"\nend if\nend program t\n",
        ["A-"]
    };

    block_if_five_tier_grade_b_plus => {
        "program t\ninteger :: pts = 88\nif (pts >= 97) then\nprint *, \"A+\"\nelse if (pts >= 93) then\nprint *, \"A\"\nelse if (pts >= 90) then\nprint *, \"A-\"\nelse if (pts >= 87) then\nprint *, \"B+\"\nelse\nprint *, \"B-\"\nend if\nend program t\n",
        ["B+"]
    };

    block_if_http_status_success => {
        "program t\ninteger :: status = 201\nif (status >= 500) then\nprint *, \"server\"\nelse if (status >= 400) then\nprint *, \"client\"\nelse if (status >= 300) then\nprint *, \"redirect\"\nelse if (status >= 200) then\nprint *, \"success\"\nelse\nprint *, \"info\"\nend if\nend program t\n",
        ["success"]
    };

    block_if_http_status_client_error => {
        "program t\ninteger :: status = 404\nif (status >= 500) then\nprint *, \"server\"\nelse if (status >= 400) then\nprint *, \"client\"\nelse if (status >= 300) then\nprint *, \"redirect\"\nelse if (status >= 200) then\nprint *, \"success\"\nelse\nprint *, \"info\"\nend if\nend program t\n",
        ["client"]
    };

    block_if_quadrant_second => {
        "program t\ninteger :: px = -3, py = 4\nif (px >= 0 .and. py >= 0) then\nprint *, \"q1\"\nelse if (px < 0 .and. py >= 0) then\nprint *, \"q2\"\nelse if (px < 0 .and. py < 0) then\nprint *, \"q3\"\nelse\nprint *, \"q4\"\nend if\nend program t\n",
        ["q2"]
    };

    block_if_month_name_december => {
        "program t\ninteger :: m = 12\nif (m == 1) then\nprint *, \"jan\"\nelse if (m == 4) then\nprint *, \"apr\"\nelse if (m == 7) then\nprint *, \"jul\"\nelse if (m == 10) then\nprint *, \"oct\"\nelse if (m == 12) then\nprint *, \"dec\"\nelse\nprint *, \"other\"\nend if\nend program t\n",
        ["dec"]
    };

    block_if_size_category_medium => {
        "program t\ninteger :: n = 42\nif (n < 10) then\nprint *, \"tiny\"\nelse if (n < 100) then\nprint *, \"small\"\nelse if (n < 1000) then\nprint *, \"medium\"\nelse if (n < 10000) then\nprint *, \"large\"\nelse\nprint *, \"huge\"\nend if\nend program t\n",
        ["small"]
    };

    block_if_dice_sum_lucky_seven => {
        "program t\ninteger :: d1 = 3, d2 = 4\nif (d1 + d2 == 2) then\nprint *, \"snake\"\nelse if (d1 + d2 == 7) then\nprint *, \"lucky\"\nelse if (d1 + d2 == 12) then\nprint *, \"box\"\nelse\nprint *, \"plain\"\nend if\nend program t\n",
        ["lucky"]
    };

    block_if_priority_level_high => {
        "program t\ninteger :: pri = 1\nif (pri == 0) then\nprint *, \"critical\"\nelse if (pri == 1) then\nprint *, \"high\"\nelse if (pri == 2) then\nprint *, \"normal\"\nelse if (pri == 3) then\nprint *, \"low\"\nelse\nprint *, \"deferred\"\nend if\nend program t\n",
        ["high"]
    };

    // ── IF without ELSE ─────────────────────────────────────────────────

    if_no_else_true_branch_runs => {
        "program t\nif (7 > 3) then\nprint *, \"ran\"\nend if\nend program t\n",
        ["ran"]
    };

    if_no_else_false_branch_skipped => {
        "program t\nif (2 > 9) then\nprint *, \"skip\"\nend if\nprint *, \"after\"\nend program t\n",
        ["after"]
    };

    if_no_else_guard_positive_value => {
        "program t\ninteger :: n = 5\nif (n > 0) then\nprint *, \"ok\"\nend if\nend program t\n",
        ["ok"]
    };

    if_no_else_guard_zero_falls_through => {
        "program t\ninteger :: n = 0\nif (n > 0) then\nprint *, \"ok\"\nend if\nprint *, \"done\"\nend program t\n",
        ["done"]
    };

    if_no_else_sequential_first_matches => {
        "program t\ninteger :: x = 4\nif (x == 4) then\nprint *, \"four\"\nend if\nif (x == 5) then\nprint *, \"five\"\nend if\nend program t\n",
        ["four"]
    };

    if_no_else_sequential_none_match => {
        "program t\ninteger :: x = 4\nif (x == 1) then\nprint *, \"one\"\nend if\nif (x == 2) then\nprint *, \"two\"\nend if\nprint *, \"end\"\nend program t\n",
        ["end"]
    };

    if_no_else_real_above_threshold => {
        "program t\nreal :: r = 3.5\nif (r > 3.0) then\nprint *, \"above\"\nend if\nend program t\n",
        ["above"]
    };

    if_no_else_char_starts_with_a => {
        "program t\ncharacter(len=5) :: word = \"alpha\"\nif (word(1:1) == \"a\") then\nprint *, \"a-word\"\nend if\nend program t\n",
        ["a-word"]
    };

    if_no_else_multi_statement_body => {
        "program t\nif (1 == 1) then\nprint *, \"step1\"\nprint *, \"step2\"\nend if\nend program t\n",
        ["step1", "step2"]
    };

    if_no_else_outer_true_inner_false => {
        "program t\nif (3 > 1) then\nif (5 < 2) then\nprint *, \"inner\"\nend if\nprint *, \"outer\"\nend if\nend program t\n",
        ["outer"]
    };

    if_nested_true_outer_false_else => {
        "program t\ninteger :: x = -4\nif (x > 0) then\nif (x > 10) then\nprint *, \"deep\"\nend if\nprint *, \"outer-true\"\nelse\nprint *, \"outer-false\"\nif (x == -4) then\nprint *, \"inner-match\"\nelse\nprint *, \"inner-miss\"\nend if\nend if\nend program t\n",
        ["outer-false", "inner-match"]
    };

    if_nested_if_with_else_branch_only => {
        "program t\ninteger :: y = 7\nif (y < 0) then\nprint *, \"negative\"\nelse\nif (y > 10) then\nprint *, \"double-digits\"\nelse if (y == 7) then\nprint *, \"lucky\"\nelse\nprint *, \"small\"\nend if\nend if\nend program t\n",
        ["lucky"]
    };

    // ── Compound logical conditions ─────────────────────────────────────

    if_compound_and_or_first_clause_true => {
        "program t\nif (1 > 0 .and. 2 > 1 .or. 0 > 5) then\nprint *, \"yes\"\nelse\nprint *, \"no\"\nend if\nend program t\n",
        ["yes"]
    };

    if_compound_not_and_both_required => {
        "program t\nif (.not. (3 > 5 .and. 1 > 2)) then\nprint *, \"pass\"\nelse\nprint *, \"fail\"\nend if\nend program t\n",
        ["pass"]
    };

    if_compound_parenthesized_or_inside_and => {
        "program t\nif ((1 > 5 .or. 2 > 1) .and. 3 > 0) then\nprint *, \"hit\"\nelse\nprint *, \"miss\"\nend if\nend program t\n",
        ["hit"]
    };

    if_compound_neqv_as_xor_true => {
        "program t\nlogical :: a, b\na = .true.\nb = .false.\nif (a .neqv. b) then\nprint *, \"xor\"\nelse\nprint *, \"same\"\nend if\nend program t\n",
        ["xor"]
    };

    if_compound_triple_and_all_true => {
        "program t\nif (1 < 2 .and. 2 < 3 .and. 3 < 4) then\nprint *, \"all\"\nelse\nprint *, \"not\"\nend if\nend program t\n",
        ["all"]
    };

    if_compound_triple_or_middle_true => {
        "program t\nif (1 > 5 .or. 2 == 2 .or. 3 > 9) then\nprint *, \"any\"\nelse\nprint *, \"none\"\nend if\nend program t\n",
        ["any"]
    };

    if_compound_de_morgan_not_or => {
        "program t\nif (.not. (1 > 5 .or. 2 > 6)) then\nprint *, \"neither\"\nelse\nprint *, \"some\"\nend if\nend program t\n",
        ["neither"]
    };

    if_compound_logical_vars_with_compare => {
        "program t\nlogical :: flag\ninteger :: n\nflag = .true.\nn = 8\nif (flag .and. n >= 5) then\nprint *, \"ready\"\nelse\nprint *, \"wait\"\nend if\nend program t\n",
        ["ready"]
    };

    if_compound_eqv_and_compare_mixed => {
        "program t\nif ((.true. .eqv. .true.) .and. 4 > 2) then\nprint *, \"joint\"\nelse\nprint *, \"split\"\nend if\nend program t\n",
        ["joint"]
    };

    if_compound_eqv_false_hits_else => {
        "program t\nif (.true. .eqv. .false.) then\nprint *, \"same\"\nelse\nprint *, \"diff\"\nend if\nend program t\n",
        ["diff"]
    };

    if_compound_neqv_true_hits_then => {
        "program t\nif (.true. .neqv. .false.) then\nprint *, \"xor\"\nelse\nprint *, \"same\"\nend if\nend program t\n",
        ["xor"]
    };

    if_compound_nested_not_and_or => {
        "program t\nif (.not. ((1 > 2) .and. (.not. (3 < 4)))) then\nprint *, \"open\"\nelse\nprint *, \"shut\"\nend if\nend program t\n",
        ["open"]
    };

    if_compound_grouped_precedence => {
        "program t\nif ((1 < 2 .and. 2 < 3) .or. (4 < 1 .and. 9 < 10)) then\nprint *, \"yes\"\nelse\nprint *, \"no\"\nend if\nend program t\n",
        ["yes"]
    };

    if_compound_precedence_false_branch => {
        "program t\nif ((1 > 2 .or. 3 > 4) .and. .not. (5 < 6)) then\nprint *, \"bad\"\nelse\nprint *, \"good\"\nend if\nend program t\n",
        ["good"]
    };

    // ── Character comparisons in IF ─────────────────────────────────────

    if_char_literal_equality => {
        "program t\nif ('fortran' == 'fortran') then\nprint *, \"match\"\nelse\nprint *, \"diff\"\nend if\nend program t\n",
        ["match"]
    };

    if_char_literal_inequality => {
        "program t\nif ('cat' /= 'dog') then\nprint *, \"distinct\"\nelse\nprint *, \"same\"\nend if\nend program t\n",
        ["distinct"]
    };

    if_char_lexicographic_less => {
        "program t\nif ('apple' < 'banana') then\nprint *, \"before\"\nelse\nprint *, \"after\"\nend if\nend program t\n",
        ["before"]
    };

    if_char_lexicographic_greater_equal => {
        "program t\nif ('zebra' >= 'yak') then\nprint *, \"gte\"\nelse\nprint *, \"lt\"\nend if\nend program t\n",
        ["gte"]
    };

    if_char_variable_equals_literal => {
        "program t\ncharacter(len=6) :: tag = \"vybe\"\nif (tag == \"vybe\") then\nprint *, \"tag-ok\"\nelse\nprint *, \"tag-bad\"\nend if\nend program t\n",
        ["tag-ok"]
    };

    if_char_index_positive_in_branch => {
        "program t\ncharacter(len=12) :: hay = \"hello world\"\nif (index(hay, \"world\") > 0) then\nprint *, \"found\"\nelse\nprint *, \"missing\"\nend if\nend program t\n",
        ["found"]
    };

    // ── Legacy arithmetic IF ────────────────────────────────────────────

    arith_if_negative_label_branch => {
        "program t\nreal :: x = -2.5\nif (x) 10, 20, 30\n10 print *, \"negative\"; goto 99\n20 print *, \"zero\"; goto 99\n30 print *, \"positive\"\n99 continue\nend program t\n",
        ["negative"]
    };

    arith_if_zero_label_branch => {
        "program t\nreal :: x = 0.0\nif (x) 10, 20, 30\n10 print *, \"negative\"; goto 99\n20 print *, \"zero\"; goto 99\n30 print *, \"positive\"\n99 continue\nend program t\n",
        ["zero"]
    };

    arith_if_positive_label_branch => {
        "program t\nreal :: x = 7.0\nif (x) 10, 20, 30\n10 print *, \"negative\"; goto 99\n20 print *, \"zero\"; goto 99\n30 print *, \"positive\"\n99 continue\nend program t\n",
        ["positive"]
    };

    arith_if_integer_expression_zero => {
        "program t\ninteger :: n = 0\nif (n) 10, 20, 30\n10 print *, \"neg\"; goto 99\n20 print *, \"zer\"; goto 99\n30 print *, \"pos\"\n99 continue\nend program t\n",
        ["zer"]
    };

    if_single_line_if_true_path => {
        "program t\nif (1 == 1) print *, 'single-true'\nprint *, 'after'\nend program t\n",
        ["single-true", "after"]
    };

    if_single_line_if_false_path => {
        "program t\nif (1 == 2) print *, 'single-false'\nprint *, 'after'\nend program t\n",
        ["after"]
    };

    if_nested_three_level_guard_chain => {
        "program t\ninteger :: x\nx = 10\nif (x > 0) then\n    if (x > 100) then\n        print *, 'big'\n    else if (x == 10) then\n        print *, 'exact-10'\n    else\n        print *, 'small-positive'\n    end if\nelse\n    print *, 'non-positive'\nend if\n",
        ["exact-10"]
    };

    if_elif_chain_with_parenthesized_boundaries => {
        "program t\ninteger :: n\nn = 4\nif ((n / 2) == 1) then\nprint *, 'one'\nelse if (n > 3 .and. n < 6) then\nprint *, 'mid'\nelse\nprint *, 'other'\nend if\nend program t\n",
        ["mid"]
    };
}
