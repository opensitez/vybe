//! Extended control-flow coverage: GOTO, STOP/END, nested CYCLE/EXIT,
//! labeled CONTINUE, construct END boundaries, and multiple exit paths.
//! Distinct from `test_control_flow.rs` (basic if/do/select only).

fortran_cases! {
    // ── GOTO (legacy labelled branches) ────────────────────────────────

    goto_skip_unreachable_assignment => {
        "program t\ninteger :: x = 0\ngoto 10\nx = 999\n10 continue\nprint *, x\nend program t\n",
        ["0"]
    };

    goto_forward_reaches_target_label => {
        "program t\ngoto 20\n10 print *, 'skip'\ngoto 30\n20 print *, 'landed'\n30 continue\nend program t\n",
        ["landed"]
    };

    goto_f77_style_loop_sum_fifteen => {
        "program t\ninteger :: i, s\ni = 1\ns = 0\n10 if (i > 5) goto 20\ns = s + i\ni = i + 1\ngoto 10\n20 print *, s\nend program t\n",
        ["15"]
    };

    computed_goto_selects_second_label => {
        "program t\ninteger :: n = 2\ngo to (10, 20, 30), n\n10 print *, 'one'; goto 99\n20 print *, 'two'; goto 99\n30 print *, 'three'\n99 continue\nend program t\n",
        ["two"]
    };

    computed_goto_index_out_of_range_falls_through => {
        "program t\ninteger :: n = 4\ngo to (10, 20), n\nprint *, 'fallthrough'\n10 print *, 'one'\n20 print *, 'two'\nend program t\n",
        ["fallthrough"]
    };

    go_to_spelling_variant_reaches_label => {
        "program t\ngo to 20\n10 print *, 'miss'\n20 print *, 'hit'\nend program t\n",
        ["hit"]
    };

    goto_conditional_branch_to_common_label => {
        "program t\ninteger :: n = 4\nif (n < 5) goto 10\nprint *, 'high'\ngoto 20\n10 print *, 'low'\n20 continue\nend program t\n",
        ["low"]
    };

    goto_inside_do_escapes_iteration_body => {
        "program t\ninteger :: i, s\ns = 0\ndo i = 1, 10\nif (i == 4) goto 30\ns = s + i\nend do\n30 print *, s\nend program t\n",
        ["6"]
    };

    goto_chain_through_intermediate_label => {
        "program t\ngoto 10\nprint *, 'start'\n10 goto 20\nprint *, 'mid'\n20 print *, 'end'\nend program t\n",
        ["end"]
    };

    goto_nested_label_after_if => {
        "program t\ninteger :: x\nx = 0\nif (x == 0) goto 10\nx = 7\n10 print *, x\nend program t\n",
        ["0"]
    };

    // ── STOP / END / RETURN ──────────────────────────────────────────────

    stop_int_halts_before_tail_print => {
        "program t\nprint *, 'before'\nstop 0\nprint *, 'after'\nend program t\n",
        ["before"]
    };

    stop_string_halts_before_tail_print => {
        "program t\nprint *, 'ready'\nstop 'done'\nprint *, 'tail'\nend program t\n",
        ["ready"]
    };

    guarded_stop_not_taken_continues => {
        "program t\nlogical :: ok = .true.\nif (.not. ok) stop 1\nprint *, 'ok'\nend program t\n",
        ["ok"]
    };

    return_ends_program_before_second_print => {
        "program t\nprint *, 'first'\nreturn\nprint *, 'second'\nend program t\n",
        ["first"]
    };

    return_inside_if_skips_trailing_code => {
        "program t\ninteger :: flag = 1\nif (flag == 1) then\nprint *, 'in'\nreturn\nend if\nprint *, 'out'\nend program t\n",
        ["in"]
    };

    return_inside_nested_do_skips_rest => {
        "program t\ninteger :: i\ninteger :: s = 0\ni = 0\ndo while (i < 5)\ni = i + 1\ns = s + i\nif (i == 2) return\nend do\nprint *, 'after'\nprint *, s\nend program t\n",
        []
    };

    stop_after_true_guard_in_if => {
        "program t\ninteger :: code = 1\nif (code /= 0) stop 0\nprint *, 'run'\nend program t\n",
        []
    };

    // ── CYCLE / EXIT in nested loops ─────────────────────────────────────

    nested_exit_inner_leaves_outer_running => {
        "program t\ninteger :: i, j, c\nc = 0\ndo i = 1, 3\ndo j = 1, 5\nif (j > 2) exit\nc = c + 1\nend do\nend do\nprint *, c\nend program t\n",
        ["6"]
    };

    nested_exit_inner_at_j_three_yields_nine => {
        "program t\ninteger :: i, j, c\nc = 0\ndo i = 1, 3\ndo j = 1, 10\nif (j > 3) exit\nc = c + 1\nend do\nend do\nprint *, c\nend program t\n",
        ["9"]
    };

    nested_cycle_inner_skips_j_equals_two => {
        "program t\ninteger :: i, j, c\nc = 0\ndo i = 1, 3\ndo j = 1, 4\nif (j == 2) cycle\nc = c + 1\nend do\nend do\nprint *, c\nend program t\n",
        ["9"]
    };

    triple_nested_exit_innermost_at_k_three => {
        "program t\ninteger :: i, j, k, c\nc = 0\ndo i = 1, 2\ndo j = 1, 2\ndo k = 1, 5\nif (k == 3) exit\nc = c + 1\nend do\nend do\nend do\nprint *, c\nend program t\n",
        ["8"]
    };

    named_exit_outer_from_inner_loop => {
        "program t\ninteger :: i, j, c\nc = 0\nouter: do i = 1, 3\ninner: do j = 1, 5\nif (j > 2) exit outer\nc = c + 1\nend do inner\nend do outer\nprint *, c\nend program t\n",
        ["6"]
    };

    named_cycle_outer_skips_rest_of_inner => {
        "program t\ninteger :: i, j, c\nc = 0\nouter: do i = 1, 3\ninner: do j = 1, 4\nif (j == 2) cycle outer\nc = c + 1\nend do inner\nend do outer\nprint *, c\nend program t\n",
        ["3"]
    };

    named_exit_middle_from_deep_loop => {
        "program t\ninteger :: i, j, k, c\nc = 0\nouter: do i = 1, 4\nmid: do j = 1, 4\ninner: do k = 1, 4\nif (j == 2 .and. k == 2) exit mid\nc = c + 1\nend do inner\nend do mid\nend do outer\nprint *, c\nend program t\n",
        ["44"]
    };

    named_cycle_middle_from_deep_loop => {
        "program t\ninteger :: i, j, k, c\nc = 0\nouter: do i = 1, 3\nmid: do j = 1, 3\ninner: do k = 1, 3\nif (k == 2) cycle mid\nc = c + 1\nend do inner\nend do mid\nend do outer\nprint *, c\nend program t\n",
        ["12"]
    };

    do_while_exit_at_fourth_iteration => {
        "program t\ninteger :: n = 0\ndo while (n < 10)\nn = n + 1\nif (n == 4) exit\nend do\nprint *, n\nend program t\n",
        ["4"]
    };

    exit_when_running_sum_exceeds_twenty => {
        "program t\ninteger :: i, s\ns = 0\ndo i = 1, 10\ns = s + i\nif (s > 20) exit\nend do\nprint *, s\nend program t\n",
        ["21"]
    };

    cycle_skip_multiples_of_five_in_range => {
        "program t\ninteger :: i, s\ns = 0\ndo i = 1, 15\nif (mod(i, 5) == 0) cycle\ns = s + i\nend do\nprint *, s\nend program t\n",
        ["90"]
    };

    nested_mix_cycle_inner_exit_outer => {
        "program t\ninteger :: i, j, s\ns = 0\nouter: do i = 1, 5\ninner: do j = 1, 5\nif (j == 1) cycle inner\nif (i == 4) exit outer\ns = s + 1\nend do inner\nend do outer\nprint *, s\nend program t\n",
        ["8"]
    };

    // ── Labeled CONTINUE (legacy DO/CONTINUE) ────────────────────────────

    labeled_continue_skip_evens_sum_odds => {
        "program t\ninteger :: i, s\ns = 0\ndo i = 1, 6\nif (mod(i, 2) == 0) goto 100\ns = s + i\n100 continue\nend do\nprint *, s\nend program t\n",
        ["9"]
    };

    labeled_do_sum_one_to_four => {
        "program t\ninteger :: i, s\ns = 0\ndo 100 i = 1, 4\ns = s + i\n100 continue\nprint *, s\nend program t\n",
        ["10"]
    };

    nested_labeled_do_counts_nine => {
        "program t\ninteger :: i, j, c\nc = 0\ndo 200 i = 1, 3\ndo 100 j = 1, 3\nc = c + 1\n100 continue\n200 continue\nprint *, c\nend program t\n",
        ["9"]
    };

    labeled_do_step_three_sum_twenty_two => {
        "program t\ninteger :: i, s\ns = 0\ndo 50 i = 1, 10, 3\ns = s + i\n50 continue\nprint *, s\nend program t\n",
        ["22"]
    };

    labeled_do_empty_range_keeps_initial => {
        "program t\ninteger :: i, s\ns = 77\ndo 100 i = 8, 3\ns = s + i\n100 continue\nprint *, s\nend program t\n",
        ["77"]
    };

    labeled_continue_after_print_in_do => {
        "program t\ninteger :: i\ndo i = 1, 4\nif (i == 3) goto 200\nprint *, i\n200 continue\nend do\nend program t\n",
        ["1", "2", "4"]
    };

    // ── Block / construct END boundaries ─────────────────────────────────

    if_then_end_if_reaches_following_print => {
        "program t\ninteger :: x = 3\nif (x > 0) then\nprint *, 'in'\nend if\nprint *, 'out'\nend program t\n",
        ["in", "out"]
    };

    nested_if_end_if_resolves_inner_else => {
        "program t\ninteger :: a = 1, b = 0, r\nif (a == 1) then\nif (b == 1) then\nr = 10\nelse\nr = 20\nend if\nelse\nr = 30\nend if\nprint *, r\nend program t\n",
        ["20"]
    };

    end_do_named_outer_loop_tag => {
        "program t\ninteger :: i, s\ns = 0\nouter: do i = 1, 4\ns = s + i\nend do outer\nprint *, s\nend program t\n",
        ["10"]
    };

    end_select_after_case_block => {
        "program t\ninteger :: v = 2, r\nselect case (v)\ncase (1)\nr = 10\ncase (2)\nr = 20\ncase default\nr = 99\nend select\nprint *, r\nend program t\n",
        ["20"]
    };

    block_end_block_local_doubles_input => {
        "program t\ninteger :: x = 6\nblock\ninteger :: y\ny = x * 2\nprint *, y\nend block\nprint *, x\nend program t\n",
        ["12", "6"]
    };

    nested_block_end_block_prints_inner => {
        "program t\ninteger :: a = 2\nblock\ninteger :: b\nb = a + 3\nblock\ninteger :: c\nc = b + 5\nprint *, c\nend block\nend block\nprint *, a\nend program t\n",
        ["10", "2"]
    };

    // ── Multiple exit paths ──────────────────────────────────────────────

    elseif_chain_hits_second_of_three => {
        "program t\ninteger :: score = 82\nif (score >= 90) then\nprint *, 'A'\nelse if (score >= 80) then\nprint *, 'B'\nelse\nprint *, 'C'\nend if\nend program t\n",
        ["B"]
    };

    elseif_chain_falls_to_else => {
        "program t\ninteger :: score = 55\nif (score >= 90) then\nprint *, 'A'\nelse if (score >= 80) then\nprint *, 'B'\nelse\nprint *, 'C'\nend if\nend program t\n",
        ["C"]
    };

    select_case_second_of_three_paths => {
        "program t\ninteger :: day = 2\nselect case (day)\ncase (1)\nprint *, 'mon'\ncase (2)\nprint *, 'tue'\ncase (3)\nprint *, 'wed'\ncase default\nprint *, 'other'\nend select\nend program t\n",
        ["tue"]
    };

    select_case_default_when_unmatched => {
        "program t\ninteger :: code = 404\nselect case (code)\ncase (200)\nprint *, 'ok'\ncase (301)\nprint *, 'move'\ncase default\nprint *, 'err'\nend select\nend program t\n",
        ["err"]
    };

    sign_classifier_negative_branch => {
        "program t\ninteger :: v = -8\nif (v > 0) then\nprint *, 'pos'\nelse if (v < 0) then\nprint *, 'neg'\nelse\nprint *, 'zero'\nend if\nend program t\n",
        ["neg"]
    };

    search_loop_exit_on_first_match => {
        "program t\ninteger :: arr(5), i, found\narr = [3, 7, 2, 9, 4]\nfound = 0\ndo i = 1, 5\nif (arr(i) == 2) then\nfound = i\nexit\nend if\nend do\nprint *, found\nend program t\n",
        ["3"]
    };

    if_guarded_exit_inside_do => {
        "program t\ninteger :: i, hits\nhits = 0\ndo i = 1, 12\nif (mod(i, 4) == 0) then\nhits = hits + 1\nif (hits == 2) exit\nend if\nend do\nprint *, i\nend program t\n",
        ["8"]
    };

    parity_cycle_or_accumulate_odds => {
        "program t\ninteger :: i, s\ns = 0\ndo i = 1, 9\nif (mod(i, 2) == 0) cycle\ns = s + i\nend do\nprint *, s\nend program t\n",
        ["25"]
    };

    nested_cycle_skips_named_inner_and_keeps_outer_count => {
        "program t\ninteger :: i, j, s\ns = 0\nouter: do i = 1, 4\ninner: do j = 1, 4\nif (mod(j, 2) == 0) cycle inner\ns = s + 1\nend do inner\nend do outer\nprint *, s\nend program t\n",
        ["8"]
    };

    tiered_threshold_picks_middle_band => {
        "program t\ninteger :: val = 45\nif (val < 20) then\nprint *, 'low'\nelse if (val < 60) then\nprint *, 'mid'\nelse\nprint *, 'high'\nend if\nend program t\n",
        ["mid"]
    };

    select_inside_do_first_matching_case => {
        "program t\ninteger :: i, hits\nhits = 0\ndo i = 1, 6\nselect case (mod(i, 3))\ncase (0)\nhits = hits + 10\ncase (1)\nhits = hits + 1\ncase (2)\nhits = hits + 2\nend select\nend do\nprint *, hits\nend program t\n",
        ["15"]
    };

    twin_conditional_exits_first_match_wins => {
        "program t\ninteger :: x = 5, y\nif (x == 5) then\ny = 1\nelse if (x == 5) then\ny = 2\nelse\ny = 3\nend if\nprint *, y\nend program t\n",
        ["1"]
    };

    loop_exit_on_target_value_found => {
        "program t\ninteger :: i\ndo i = 10, 99\nif (mod(i, 13) == 0 .and. mod(i, 7) == 0) exit\nend do\nprint *, i\nend program t\n",
        ["91"]
    };

    loop_exit_on_named_outer_target_hit => {
        "program t\ninteger :: i, j\nouter_loop: do i = 1, 4\ninner_loop: do j = 1, 6\nif (i == 3 .and. j == 2) exit outer_loop\nend do inner_loop\nend do outer_loop\nprint *, i\nprint *, j\nend program t\n",
        ["3", "1"]
    };

    assigned_goto_selects_assigned_label => {
        "program t\ninteger :: n, x\nx = 0\nassign 20 to n\ngo to n\nx = 99\n20 x = 20\nprint *, x\nend program t\n",
        ["20"]
    };

    arithmetic_if_positive_label_1 => {
        "program t\ninteger :: x, y\nx = 5\ny = 0\nif (x) 10,20,30\n10 y = 10\ngoto 40\n20 y = 20\ngoto 40\n30 y = 30\n40 continue\nprint *, y\nend program t\n",
        ["10"]
    };

    arithmetic_if_zero_label_1 => {
        "program t\ninteger :: x, y\nx = 0\ny = 0\nif (x) 10,20,30\n10 y = 10\ngoto 40\n20 y = 20\ngoto 40\n30 y = 30\n40 continue\nprint *, y\nend program t\n",
        ["20"]
    };

    arithmetic_if_negative_label_1 => {
        "program t\ninteger :: x, y\nx = -1\ny = 0\nif (x) 10,20,30\n10 y = 10\ngoto 40\n20 y = 20\ngoto 40\n30 y = 30\n40 continue\nprint *, y\nend program t\n",
        ["10"]
    };

    select_type_integer_allocation_prints_int => {
        "program t\nclass(*), allocatable :: x\nallocate(integer::x)\nselect type(x)\n type is(integer)\n  print *, 'int'\n class default\n  print *, 'other'\nend select\nend program t\n",
        ["int"]
    };

    select_type_character_defaults_to_class_default => {
        "program t\nclass(*), allocatable :: x\nallocate(character(len=4)::x)\nselect type(x)\n type is(integer)\n  print *, 'int'\n type is(real)\n  print *, 'real'\n class default\n  print *, 'other'\nend select\nend program t\n",
        ["other"]
    };

    select_case_range_branches => {
        "program t\ninteger :: day = 15\nselect case (day)\ncase (1:7)\n print *, 'week'\ncase (8:14)\n print *, 'half'\ncase (15:21)\n print *, 'three'\ncase default\n print *, 'other'\nend select\nend program t\n",
        ["three"]
    };

    where_elsewhere_masked_update => {
        "program t\ninteger :: a(5), b(5)\na = 0\nb = 1\nwhere (mod(a, 2) == 0)\n  b = b + 10\nelsewhere (a < 0)\n  b = 99\nelsewhere\n  b = 7\nend where\nprint *, b(1)\nprint *, b(3)\nprint *, b(5)\nend program t\n",
        ["11", "11", "11"]
    };
}
