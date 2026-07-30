//! Extended named DO loop coverage: EXIT/CYCLE targeting outer/mid/inner,
//! triple-nested named loops, and mixed named/unnamed constructs.
//! Distinct from `test_named_loops.rs` and `test_control_flow_extended.rs`.

fortran_cases! {
    // ── EXIT outer from nested named loops ─────────────────────────────

    exit_outer_when_ij_product_reaches_nine => {
        "program t\ninteger :: i, j\nouter: do i = 1, 5\ninner: do j = 1, 5\nif (i * j >= 9) exit outer\nend do inner\nend do outer\nprint *, i\nprint *, j\nend program t\n",
        ["2", "5"]
    };

    exit_outer_from_deep_when_sum_is_ten => {
        "program t\ninteger :: i, j, k\nouter: do i = 1, 6\nmid: do j = 1, 6\ninner: do k = 1, 6\nif (i + j + k == 10) exit outer\nend do inner\nend do mid\nend do outer\nprint *, i\nprint *, j\nprint *, k\nend program t\n",
        ["1", "3", "6"]
    };

    exit_outer_on_first_outer_iteration => {
        "program t\ninteger :: i, j\ngo: do i = 1, 10\ninner: do j = 1, 10\nif (i == 1 .and. j == 1) exit go\nend do inner\nend do go\nprint *, i\nprint *, j\nend program t\n",
        ["1", "1"]
    };

    exit_outer_when_i_equals_three_j_equals_two => {
        "program t\ninteger :: i, j\nouter: do i = 1, 5\ninner: do j = 1, 5\nif (i == 3 .and. j == 2) exit outer\nend do inner\nend do outer\nprint *, i\nprint *, j\nend program t\n",
        ["3", "2"]
    };

    exit_outer_from_triple_when_product_is_twentyfour => {
        "program t\ninteger :: i, j, k\nfound: do i = 1, 6\nmid: do j = 1, 6\ndeep: do k = 1, 6\nif (i * j * k == 24) exit found\nend do deep\nend do mid\nend do found\nprint *, i * j * k\nend program t\n",
        ["24"]
    };

    exit_outer_leaves_mid_and_inner => {
        "program t\ninteger :: i, j, k, c\nc = 0\nouter: do i = 1, 4\nmid: do j = 1, 4\ninner: do k = 1, 4\nif (i == 2 .and. j == 2 .and. k == 1) exit outer\nc = c + 1\nend do inner\nend do mid\nend do outer\nprint *, c\nend program t\n",
        ["21"]
    };

    exit_outer_from_named_inner_unnamed_mid => {
        "program t\ninteger :: i, j, k\nouter: do i = 1, 5\ndo j = 1, 5\ninner: do k = 1, 5\nif (i + j + k == 8) exit outer\nend do inner\nend do\nend do outer\nprint *, i\nprint *, j\nprint *, k\nend program t\n",
        ["1", "2", "5"]
    };

    exit_outer_when_running_count_hits_five => {
        "program t\ninteger :: i, j, c\nc = 0\nouter: do i = 1, 10\ninner: do j = 1, 10\nc = c + 1\nif (c == 5) exit outer\nend do inner\nend do outer\nprint *, c\nprint *, i\nprint *, j\nend program t\n",
        ["5", "1", "5"]
    };

    exit_outer_from_do_while_named => {
        "program t\ninteger :: n = 0, j\nspin: do while (n < 20)\nn = n + 1\ninner: do j = 1, 5\nif (n == 7) exit spin\nend do inner\nend do spin\nprint *, n\nend program t\n",
        ["7"]
    };

    exit_outer_skips_remaining_outers => {
        "program t\ninteger :: i, j, total\ntotal = 0\nouter: do i = 1, 5\ninner: do j = 1, 5\ntotal = total + 1\nif (i == 2 .and. j == 3) exit outer\nend do inner\nend do outer\nprint *, total\nend program t\n",
        ["8"]
    };

    // ── EXIT mid from triple nested ────────────────────────────────────

    exit_mid_when_jk_sum_is_five => {
        "program t\ninteger :: i, j, k, c\nc = 0\nouter: do i = 1, 3\nmid: do j = 1, 4\ninner: do k = 1, 4\nif (j + k == 5) exit mid\nc = c + 1\nend do inner\nend do mid\nend do outer\nprint *, c\nend program t\n",
        ["24"]
    };

    exit_mid_on_j_three_k_one => {
        "program t\ninteger :: i, j, k, c\nc = 0\nouter: do i = 1, 3\nmid: do j = 1, 5\ninner: do k = 1, 5\nif (j == 3 .and. k == 1) exit mid\nc = c + 1\nend do inner\nend do mid\nend do outer\nprint *, c\nend program t\n",
        ["30"]
    };

    exit_mid_preserves_outer_index => {
        "program t\ninteger :: i, j, k\nouter: do i = 1, 4\nmid: do j = 1, 4\ninner: do k = 1, 4\nif (j == 1 .and. k == 4) exit mid\nend do inner\nend do mid\nend do outer\nprint *, i\nprint *, j\nprint *, k\nend program t\n",
        ["4", "1", "4"]
    };

    exit_mid_from_unnamed_inner_named_mid => {
        "program t\ninteger :: i, j, k, c\nc = 0\nouter: do i = 1, 2\ndo j = 1, 3\nmid: do k = 1, 3\nif (k == 2) exit mid\nc = c + 1\nend do mid\nend do\nend do outer\nprint *, c\nend program t\n",
        ["6"]
    };

    exit_mid_when_j_times_k_is_six => {
        "program t\ninteger :: i, j, k\nouter: do i = 1, 5\nmid: do j = 1, 5\ninner: do k = 1, 5\nif (j * k == 6) exit mid\nend do inner\nend do mid\nend do outer\nprint *, i\nprint *, j\nprint *, k\nend program t\n",
        ["5", "2", "3"]
    };

    exit_mid_at_j_two_k_three_count => {
        "program t\ninteger :: i, j, k, c\nc = 0\nouter: do i = 1, 2\nmid: do j = 1, 3\ninner: do k = 1, 3\nif (j == 2 .and. k == 3) exit mid\nc = c + 1\nend do inner\nend do mid\nend do outer\nprint *, c\nend program t\n",
        ["16"]
    };

    // ── EXIT inner from nested named loops ─────────────────────────────

    exit_inner_at_j_greater_than_three => {
        "program t\ninteger :: i, j, c\nc = 0\nouter: do i = 1, 4\ninner: do j = 1, 8\nif (j > 3) exit inner\nc = c + 1\nend do inner\nend do outer\nprint *, c\nend program t\n",
        ["12"]
    };

    exit_inner_at_k_equals_three_count => {
        "program t\ninteger :: i, j, k, c\nc = 0\nouter: do i = 1, 2\nmid: do j = 1, 2\ninner: do k = 1, 5\nif (k == 3) exit inner\nc = c + 1\nend do inner\nend do mid\nend do outer\nprint *, c\nend program t\n",
        ["8"]
    };

    exit_inner_allows_outer_to_continue => {
        "program t\ninteger :: i, j, s\ns = 0\nouter: do i = 1, 3\ninner: do j = 1, 10\nif (j > 3) exit inner\ns = s + j\nend do inner\nend do outer\nprint *, s\nend program t\n",
        ["18"]
    };

    exit_inner_from_mixed_named_outer => {
        "program t\ninteger :: i, j, c\nc = 0\nouter: do i = 1, 3\ndo j = 1, 6\nif (j == 4) exit\nif (j > 4) exit outer\nc = c + 1\nend do\nend do outer\nprint *, c\nend program t\n",
        ["9"]
    };

    exit_inner_named_from_triple => {
        "program t\ninteger :: i, j, k, c\nc = 0\nouter: do i = 1, 2\nmid: do j = 1, 2\ninner: do k = 1, 6\nif (k == 4) exit inner\nc = c + 1\nend do inner\nend do mid\nend do outer\nprint *, c\nend program t\n",
        ["12"]
    };

    // ── CYCLE outer ────────────────────────────────────────────────────

    cycle_outer_when_j_equals_three => {
        "program t\ninteger :: i, j, c\nc = 0\nouter: do i = 1, 4\ninner: do j = 1, 5\nif (j == 3) cycle outer\nc = c + 1\nend do inner\nend do outer\nprint *, c\nend program t\n",
        ["8"]
    };

    cycle_outer_on_even_j => {
        "program t\ninteger :: i, j, c\nc = 0\nouter: do i = 1, 3\ninner: do j = 1, 6\nif (mod(j, 2) == 0) cycle outer\nc = c + 1\nend do inner\nend do outer\nprint *, c\nend program t\n",
        ["9"]
    };

    cycle_outer_from_triple_when_k_is_one => {
        "program t\ninteger :: i, j, k, c\nc = 0\nouter: do i = 1, 2\nmid: do j = 1, 3\ninner: do k = 1, 3\nif (k == 1) cycle outer\nc = c + 1\nend do inner\nend do mid\nend do outer\nprint *, c\nend program t\n",
        ["0"]
    };

    cycle_outer_skips_j_four_and_five => {
        "program t\ninteger :: i, j, s\ns = 0\nouter: do i = 1, 3\ninner: do j = 1, 5\nif (j == 2) cycle outer\ns = s + j\nend do inner\nend do outer\nprint *, s\nend program t\n",
        ["3"]
    };

    cycle_outer_when_i_plus_j_gt_five => {
        "program t\ninteger :: i, j, c\nc = 0\nouter: do i = 1, 4\ninner: do j = 1, 4\nif (i + j > 5) cycle outer\nc = c + 1\nend do inner\nend do outer\nprint *, c\nend program t\n",
        ["10"]
    };

    cycle_outer_on_j_equals_four => {
        "program t\ninteger :: i, j, c\nc = 0\nouter: do i = 1, 5\ninner: do j = 1, 6\nif (j == 4) cycle outer\nc = c + 1\nend do inner\nend do outer\nprint *, c\nend program t\n",
        ["15"]
    };

    // ── CYCLE mid ──────────────────────────────────────────────────────

    cycle_mid_when_k_equals_one => {
        "program t\ninteger :: i, j, k, c\nc = 0\nouter: do i = 1, 2\nmid: do j = 1, 3\ninner: do k = 1, 3\nif (k == 1) cycle mid\nc = c + 1\nend do inner\nend do mid\nend do outer\nprint *, c\nend program t\n",
        ["12"]
    };

    cycle_mid_on_k_even => {
        "program t\ninteger :: i, j, k, c\nc = 0\nouter: do i = 1, 2\nmid: do j = 1, 2\ninner: do k = 1, 4\nif (mod(k, 2) == 0) cycle mid\nc = c + 1\nend do inner\nend do mid\nend do outer\nprint *, c\nend program t\n",
        ["8"]
    };

    cycle_mid_from_deep_preserves_outer => {
        "program t\ninteger :: i, j, k, total\ntotal = 0\nouter: do i = 1, 3\nmid: do j = 1, 3\ninner: do k = 1, 3\nif (j == 2 .and. k == 1) cycle mid\ntotal = total + 1\nend do inner\nend do mid\nend do outer\nprint *, total\nend program t\n",
        ["24"]
    };

    cycle_mid_when_jk_product_is_four => {
        "program t\ninteger :: i, j, k, c\nc = 0\nouter: do i = 1, 2\nmid: do j = 1, 3\ninner: do k = 1, 3\nif (j * k == 4) cycle mid\nc = c + 1\nend do inner\nend do mid\nend do outer\nprint *, c\nend program t\n",
        ["14"]
    };

    // ── CYCLE inner ────────────────────────────────────────────────────

    cycle_inner_skip_odd_j => {
        "program t\ninteger :: i, j, c\nc = 0\nouter: do i = 1, 3\ninner: do j = 1, 6\nif (mod(j, 2) == 1) cycle inner\nc = c + 1\nend do inner\nend do outer\nprint *, c\nend program t\n",
        ["9"]
    };

    cycle_inner_when_j_equals_two => {
        "program t\ninteger :: i, j, c\nc = 0\nouter: do i = 1, 4\ninner: do j = 1, 5\nif (j == 2) cycle inner\nc = c + 1\nend do inner\nend do outer\nprint *, c\nend program t\n",
        ["16"]
    };

    cycle_inner_at_k_equals_one_in_triple => {
        "program t\ninteger :: i, j, k, c\nc = 0\nouter: do i = 1, 2\nmid: do j = 1, 2\ninner: do k = 1, 4\nif (k == 1) cycle inner\nc = c + 1\nend do inner\nend do mid\nend do outer\nprint *, c\nend program t\n",
        ["12"]
    };

    cycle_inner_preserves_outer_accumulator => {
        "program t\ninteger :: i, j, sum\nsum = 0\nouter: do i = 1, 3\ninner: do j = 1, 4\nif (j == 1) cycle inner\nsum = sum + i * j\nend do inner\nend do outer\nprint *, sum\nend program t\n",
        ["54"]
    };

    // ── Triple nested named loops ──────────────────────────────────────

    triple_named_sum_all_iterations => {
        "program t\ninteger :: i, j, k, c\nc = 0\nouter: do i = 1, 2\nmid: do j = 1, 2\ninner: do k = 1, 2\nc = c + 1\nend do inner\nend do mid\nend do outer\nprint *, c\nend program t\n",
        ["8"]
    };

    triple_named_exit_inner_count => {
        "program t\ninteger :: i, j, k, c\nc = 0\nouter: do i = 1, 2\nmid: do j = 1, 2\ninner: do k = 1, 4\nif (k == 2) exit inner\nc = c + 1\nend do inner\nend do mid\nend do outer\nprint *, c\nend program t\n",
        ["8"]
    };

    triple_named_exit_mid_at_j_two_k_three => {
        "program t\ninteger :: i, j, k, c\nc = 0\nouter: do i = 1, 3\nmid: do j = 1, 3\ninner: do k = 1, 3\nif (j == 2 .and. k == 3) exit mid\nc = c + 1\nend do inner\nend do mid\nend do outer\nprint *, c\nend program t\n",
        ["24"]
    };

    triple_named_product_scan_exit_outer => {
        "program t\ninteger :: i, j, k\nscan: do i = 1, 5\nmid: do j = 1, 5\ninner: do k = 1, 5\nif (i + j + k == 7) exit scan\nend do inner\nend do mid\nend do scan\nprint *, i\nprint *, j\nprint *, k\nend program t\n",
        ["1", "1", "5"]
    };

    triple_named_accumulate_with_cycle_mid => {
        "program t\ninteger :: i, j, k, s\ns = 0\nouter: do i = 1, 2\nmid: do j = 1, 3\ninner: do k = 1, 3\nif (k == 2) cycle mid\ns = s + 1\nend do inner\nend do mid\nend do outer\nprint *, s\nend program t\n",
        ["12"]
    };

    // ── Mixed named and unnamed loops ────────────────────────────────

    mixed_unnamed_outer_named_inner_exit_inner => {
        "program t\ninteger :: i, j, c\nc = 0\ndo i = 1, 3\ninner: do j = 1, 6\nif (j > 3) exit inner\nc = c + 1\nend do inner\nend do\nprint *, c\nend program t\n",
        ["9"]
    };

    mixed_named_outer_unnamed_inner_exit_named => {
        "program t\ninteger :: i, j\nouter: do i = 1, 5\ndo j = 1, 5\nif (j == 3) exit outer\nend do\nend do outer\nprint *, i\nprint *, j\nend program t\n",
        ["1", "3"]
    };

    mixed_named_mid_unnamed_inner_outer => {
        "program t\ninteger :: i, j, k, c\nc = 0\nouter: do i = 1, 2\ndo j = 1, 2\nmid: do k = 1, 3\nif (k == 2) exit mid\nc = c + 1\nend do mid\nend do\nend do outer\nprint *, c\nend program t\n",
        ["4"]
    };

    mixed_unnamed_exit_vs_named_exit => {
        "program t\ninteger :: i, j, c\nc = 0\nnamed: do i = 1, 4\ndo j = 1, 4\nif (j == 2) exit\nc = c + 1\nif (j == 3) exit named\nend do\nend do named\nprint *, c\nend program t\n",
        ["4"]
    };

    mixed_cycle_named_from_unnamed_inner => {
        "program t\ninteger :: i, j, c\nc = 0\nouter: do i = 1, 3\ndo j = 1, 5\nif (j == 4) cycle outer\nc = c + 1\nend do\nend do outer\nprint *, c\nend program t\n",
        ["9"]
    };

    mixed_named_inner_cycle_unnamed_outer_runs => {
        "program t\ninteger :: i, j, total\ntotal = 0\ndo i = 1, 2\ninner: do j = 1, 4\nif (j == 2) cycle inner\ntotal = total + 1\nend do inner\nend do\nprint *, total\nend program t\n",
        ["6"]
    };

    mixed_triple_only_mid_named => {
        "program t\ninteger :: i, j, k, c\nc = 0\ndo i = 1, 2\ndo j = 1, 2\nmid: do k = 1, 3\nif (k == 1) cycle mid\nc = c + 1\nend do mid\nend do\nend do\nprint *, c\nend program t\n",
        ["8"]
    };

    mixed_named_do_while_with_unnamed_inner => {
        "program t\ninteger :: n, j, c\nn = 0\nc = 0\nspin: do while (n < 5)\nn = n + 1\ndo j = 1, 3\nif (j == 2) cycle spin\nc = c + 1\nend do\nend do spin\nprint *, c\nend program t\n",
        ["5"]
    };

    mixed_exit_and_cycle_named_outer => {
        "program t\ninteger :: i, j, s\ns = 0\nouter: do i = 1, 6\ninner: do j = 1, 6\nif (j == 1) cycle inner\nif (j == 5) cycle outer\nif (i == 5 .and. j == 3) exit outer\ns = s + 1\nend do inner\nend do outer\nprint *, s\nend program t\n",
        ["17"]
    };

    mixed_only_inner_named_cycle_inner => {
        "program t\ninteger :: i, j, c\nc = 0\ndo i = 1, 3\ndo j = 1, 5\ninner: do k = 1, 1\nif (j == 3) cycle inner\nc = c + 1\nend do inner\nend do\nend do\nprint *, c\nend program t\n",
        ["12"]
    };

    named_loop_negative_step_with_exit => {
        "program t\ninteger :: i, s\ns = 0\ndown: do i = 10, 1, -1\ns = s + i\nif (i == 6) exit down\nend do down\nprint *, s\nend program t\n",
        ["40"]
    };

    named_do_while_cycles_before_sum => {
        "program t\ninteger :: n, total\nn = 0\ntotal = 0\ncount: do while (n < 6)\nn = n + 1\nif (mod(n, 2) == 0) cycle count\ntotal = total + n\nend do count\nprint *, total\nend program t\n",
        ["9"]
    };

    named_do_while_exit_stops_at_target => {
        "program t\ninteger :: n, total\nn = 0\ntotal = 0\nwatch: do while (n < 10)\nn = n + 1\nif (n == 4) exit watch\ntotal = total + n\nend do watch\nprint *, n\nprint *, total\nend program t\n",
        ["4", "6"]
    };

    named_do_while_nested_exit_outer => {
        "program t\ninteger :: outer_i, inner_i, total\nouter_i = 0\ntotal = 0\nouter_loop: do while (outer_i < 5)\nouter_i = outer_i + 1\ninner: do inner_i = 1, 10\nif (inner_i == 3 .and. outer_i == 2) exit outer_loop\ntotal = total + 1\nend do inner\nend do outer_loop\nprint *, outer_i\nprint *, total\nend program t\n",
        ["2", "4"]
    };
}
