//! Extended DO loop coverage: negative step, large step, empty ranges,
//! DO CONCURRENT, nested accumulation, exit/cycle at specific counts, labeled DO.

fortran_cases! {
    // ── Negative step ────────────────────────────────────────────────

    do_neg_step_sum_10_to_1 => {
        "program t\ninteger :: i, s\ns = 0\ndo i = 10, 1, -1\ns = s + i\nend do\nprint *, s\nend program t\n",
        ["55"]
    };

    do_neg_step_by_2_sum_10_to_1 => {
        "program t\ninteger :: i, s\ns = 0\ndo i = 10, 1, -2\ns = s + i\nend do\nprint *, s\nend program t\n",
        ["30"]
    };

    do_neg_step_sum_5_to_1 => {
        "program t\ninteger :: i, s\ns = 0\ndo i = 5, 1, -1\ns = s + i\nend do\nprint *, s\nend program t\n",
        ["15"]
    };

    do_neg_step_sum_10_to_5 => {
        "program t\ninteger :: i, s\ns = 0\ndo i = 10, 5, -1\ns = s + i\nend do\nprint *, s\nend program t\n",
        ["45"]
    };

    do_neg_step_sum_0_to_neg5 => {
        "program t\ninteger :: i, s\ns = 0\ndo i = 0, -5, -1\ns = s + i\nend do\nprint *, s\nend program t\n",
        ["-15"]
    };

    do_neg_step_by_2_sum_20_to_10 => {
        "program t\ninteger :: i, s\ns = 0\ndo i = 20, 10, -2\ns = s + i\nend do\nprint *, s\nend program t\n",
        ["90"]
    };

    do_neg_step_print_10_down_to_7 => {
        "program t\ninteger :: i\ndo i = 10, 7, -1\nprint *, i\nend do\nend program t\n",
        ["10", "9", "8", "7"]
    };

    do_neg_step_count_iterations => {
        "program t\ninteger :: i, c\nc = 0\ndo i = 10, 1, -1\nc = c + 1\nend do\nprint *, c\nend program t\n",
        ["10"]
    };

    // ── Step greater than 1 ────────────────────────────────────────────

    do_step_3_sum_1_to_10 => {
        "program t\ninteger :: i, s\ns = 0\ndo i = 1, 10, 3\ns = s + i\nend do\nprint *, s\nend program t\n",
        ["22"]
    };

    do_step_5_sum_1_to_15 => {
        "program t\ninteger :: i, s\ns = 0\ndo i = 1, 15, 5\ns = s + i\nend do\nprint *, s\nend program t\n",
        ["18"]
    };

    do_step_3_sum_2_to_20 => {
        "program t\ninteger :: i, s\ns = 0\ndo i = 2, 20, 3\ns = s + i\nend do\nprint *, s\nend program t\n",
        ["77"]
    };

    do_step_5_sum_0_to_15 => {
        "program t\ninteger :: i, s\ns = 0\ndo i = 0, 15, 5\ns = s + i\nend do\nprint *, s\nend program t\n",
        ["30"]
    };

    do_step_5_sum_1_to_11 => {
        "program t\ninteger :: i, s\ns = 0\ndo i = 1, 11, 5\ns = s + i\nend do\nprint *, s\nend program t\n",
        ["18"]
    };

    do_step_4_sum_1_to_13 => {
        "program t\ninteger :: i, s\ns = 0\ndo i = 1, 13, 4\ns = s + i\nend do\nprint *, s\nend program t\n",
        ["28"]
    };

    do_step_2_sum_1_to_9 => {
        "program t\ninteger :: i, s\ns = 0\ndo i = 1, 9, 2\ns = s + i\nend do\nprint *, s\nend program t\n",
        ["25"]
    };

    do_step_3_print_each_1_to_10 => {
        "program t\ninteger :: i\ndo i = 1, 10, 3\nprint *, i\nend do\nend program t\n",
        ["1", "4", "7", "10"]
    };

    // ── Empty DO range (start > end, positive step) ────────────────────

    do_empty_range_5_to_1_sum_zero => {
        "program t\ninteger :: i, s\ns = 0\ndo i = 5, 1\ns = s + i\nend do\nprint *, s\nend program t\n",
        ["0"]
    };

    do_empty_range_10_to_5_sum_zero => {
        "program t\ninteger :: i, s\ns = 0\ndo i = 10, 5\ns = s + i\nend do\nprint *, s\nend program t\n",
        ["0"]
    };

    do_empty_range_1_to_0_sum_zero => {
        "program t\ninteger :: i, s\ns = 0\ndo i = 1, 0\ns = s + i\nend do\nprint *, s\nend program t\n",
        ["0"]
    };

    do_empty_range_7_to_3_sum_zero => {
        "program t\ninteger :: i, s\ns = 0\ndo i = 7, 3\ns = s + i\nend do\nprint *, s\nend program t\n",
        ["0"]
    };

    do_empty_range_5_to_1_step_2 => {
        "program t\ninteger :: i, s\ns = 0\ndo i = 5, 1, 2\ns = s + i\nend do\nprint *, s\nend program t\n",
        ["0"]
    };

    do_empty_range_then_print_done => {
        "program t\ninteger :: i\ndo i = 8, 2\nprint *, i\nend do\nprint *, 'done'\nend program t\n",
        ["done"]
    };

    // ── DO CONCURRENT simple cases ─────────────────────────────────────

    do_concurrent_fill_and_read => {
        "program t\ninteger :: a(10)\ndo concurrent (i = 1:10)\na(i) = i * i\nend do\nprint *, a(4)\nend program t\n",
        ["16"]
    };

    do_concurrent_stride_2_fill => {
        "program t\ninteger :: a(10)\na = 0\ndo concurrent (i = 1:10:2)\na(i) = i\nend do\nprint *, a(5)\nend program t\n",
        ["5"]
    };

    do_concurrent_stride_2_skip_even => {
        "program t\ninteger :: a(10)\na = 0\ndo concurrent (i = 1:10:2)\na(i) = i\nend do\nprint *, a(2)\nend program t\n",
        ["0"]
    };

    do_concurrent_multiply_factor => {
        "program t\ninteger :: a(6)\ndo concurrent (i = 1:6)\na(i) = i * 3\nend do\nprint *, a(3)\nend program t\n",
        ["9"]
    };

    do_concurrent_2d_off_diagonal => {
        "program t\ninteger :: m(3,3)\nm = 0\ndo concurrent (i = 1:3, j = 1:3)\nif (i /= j) m(i,j) = i + j\nend do\nprint *, m(1,2)\nend program t\n",
        ["3"]
    };

    do_concurrent_sum_manual => {
        "program t\ninteger :: a(5), s, k\na = 0\ndo concurrent (i = 1:5)\na(i) = i\nend do\ns = 0\ndo k = 1, 5\ns = s + a(k)\nend do\nprint *, s\nend program t\n",
        ["15"]
    };

    // ── Nested DO with accumulation ────────────────────────────────────

    nested_do_sum_ij_3_by_4 => {
        "program t\ninteger :: i, j, s\ns = 0\ndo i = 1, 3\ndo j = 1, 4\ns = s + i * j\nend do\nend do\nprint *, s\nend program t\n",
        ["60"]
    };

    nested_do_sum_i_plus_j => {
        "program t\ninteger :: i, j, s\ns = 0\ndo i = 1, 2\ndo j = 1, 3\ns = s + i + j\nend do\nend do\nprint *, s\nend program t\n",
        ["21"]
    };

    nested_do_triple_count => {
        "program t\ninteger :: i, j, k, c\nc = 0\ndo i = 1, 2\ndo j = 1, 2\ndo k = 1, 2\nc = c + 1\nend do\nend do\nend do\nprint *, c\nend program t\n",
        ["8"]
    };

    nested_do_outer_accum_only => {
        "program t\ninteger :: i, j, s\ns = 0\ndo i = 1, 4\ndo j = 1, i\ns = s + 1\nend do\nend do\nprint *, s\nend program t\n",
        ["10"]
    };

    nested_do_product_table => {
        "program t\ninteger :: i, j, p\np = 1\ndo i = 1, 3\ndo j = 1, 3\np = p + i * j\nend do\nend do\nprint *, p\nend program t\n",
        ["46"]
    };

    nested_do_2_by_5_count => {
        "program t\ninteger :: i, j, c\nc = 0\ndo i = 1, 2\ndo j = 1, 5\nc = c + 1\nend do\nend do\nprint *, c\nend program t\n",
        ["10"]
    };

    // ── EXIT / CYCLE at specific iteration counts ──────────────────────

    do_exit_at_iteration_7 => {
        "program t\ninteger :: i\ndo i = 1, 20\nif (i == 7) exit\nprint *, i\nend do\nend program t\n",
        ["1", "2", "3", "4", "5", "6"]
    };

    do_exit_at_first_iteration => {
        "program t\ninteger :: i, s\ns = 0\ndo i = 1, 10\nif (i == 1) exit\ns = s + i\nend do\nprint *, s\nend program t\n",
        ["0"]
    };

    do_cycle_skip_multiples_of_3 => {
        "program t\ninteger :: i, s\ns = 0\ndo i = 1, 10\nif (mod(i, 3) == 0) cycle\ns = s + i\nend do\nprint *, s\nend program t\n",
        ["37"]
    };

    do_cycle_skip_evens_sum_odds => {
        "program t\ninteger :: i, s\ns = 0\ndo i = 1, 10\nif (mod(i, 2) == 0) cycle\ns = s + i\nend do\nprint *, s\nend program t\n",
        ["25"]
    };

    do_exit_after_five_prints => {
        "program t\ninteger :: i, c\nc = 0\ndo i = 1, 100\nif (c == 5) exit\nc = c + 1\nprint *, i\nend do\nend program t\n",
        ["1", "2", "3", "4", "5"]
    };

    do_cycle_at_2_and_5 => {
        "program t\ninteger :: i, s\ns = 0\ndo i = 1, 8\nif (i == 2 .or. i == 5) cycle\ns = s + i\nend do\nprint *, s\nend program t\n",
        ["27"]
    };

    do_exit_when_sum_exceeds_20 => {
        "program t\ninteger :: i, s\ns = 0\ndo i = 1, 10\ns = s + i\nif (s > 20) exit\nend do\nprint *, s\nend program t\n",
        ["21"]
    };

    // ── Labeled DO loops ───────────────────────────────────────────────

    labeled_do_sum_1_to_4 => {
        "program t\ninteger :: i, s\ns = 0\ndo 100 i = 1, 4\ns = s + i\n100 continue\nprint *, s\nend program t\n",
        ["10"]
    };

    labeled_do_neg_step_sum => {
        "program t\ninteger :: i, s\ns = 0\ndo 10 i = 10, 1, -1\ns = s + i\n10 continue\nprint *, s\nend program t\n",
        ["55"]
    };

    labeled_do_nested_count_3_by_3 => {
        "program t\ninteger :: i, j, c\nc = 0\ndo 200 i = 1, 3\ndo 100 j = 1, 3\nc = c + 1\n100 continue\n200 continue\nprint *, c\nend program t\n",
        ["9"]
    };

    labeled_do_step_3_sum => {
        "program t\ninteger :: i, s\ns = 0\ndo 50 i = 1, 10, 3\ns = s + i\n50 continue\nprint *, s\nend program t\n",
        ["22"]
    };

    labeled_do_empty_range => {
        "program t\ninteger :: i, s\ns = 99\ndo 100 i = 6, 2\ns = s + i\n100 continue\nprint *, s\nend program t\n",
        ["99"]
    };

    // ── Edge cases and mixed behaviors ─────────────────────────────────

    do_single_iteration_1_to_1 => {
        "program t\ninteger :: i, s\ns = 0\ndo i = 1, 1\ns = s + i\nend do\nprint *, s\nend program t\n",
        ["1"]
    };

    do_large_step_two_iterations => {
        "program t\ninteger :: i, s\ns = 0\ndo i = 1, 100, 50\ns = s + i\nend do\nprint *, s\nend program t\n",
        ["52"]
    };

    do_neg_step_exit_at_5 => {
        "program t\ninteger :: i, s\ns = 0\ndo i = 10, 1, -1\nif (i == 5) exit\ns = s + i\nend do\nprint *, s\nend program t\n",
        ["30"]
    };

    do_step_3_cycle_at_4 => {
        "program t\ninteger :: i, s\ns = 0\ndo i = 1, 13, 3\nif (i == 4) cycle\ns = s + i\nend do\nprint *, s\nend program t\n",
        ["28"]
    };
}
