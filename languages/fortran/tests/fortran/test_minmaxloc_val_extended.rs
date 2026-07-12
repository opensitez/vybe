//! Extended MINVAL/MAXVAL/MINLOC/MAXLOC with DIM=, MASK=, real arrays, and
//! known index positions. Distinct from `test_array_locators.rs` (1D locators,
//! findloc) and `test_arrays_dim_mask.rs` compile-only dim/mask smokes.

fortran_cases! {
    // ── Whole-array MINVAL/MAXVAL integers (8) ────────────────────────

    maxval_int_descending_sequence => {
        "program t\ninteger :: a(5) = [9, 7, 5, 3, 1]\nprint *, maxval(a)\nend program t\n",
        ["9"]
    };
    minval_int_ascending_sequence => {
        "program t\ninteger :: a(5) = [1, 2, 3, 4, 5]\nprint *, minval(a)\nend program t\n",
        ["1"]
    };
    maxval_int_with_negatives => {
        "program t\ninteger :: a(5) = [-10, -3, -7, -1, -5]\nprint *, maxval(a)\nend program t\n",
        ["-1"]
    };
    minval_int_with_negatives => {
        "program t\ninteger :: a(5) = [-10, -3, -7, -1, -5]\nprint *, minval(a)\nend program t\n",
        ["-10"]
    };
    maxval_int_plateau_at_end => {
        "program t\ninteger :: a(6) = [1, 2, 3, 8, 8, 8]\nprint *, maxval(a)\nend program t\n",
        ["8"]
    };
    minval_int_plateau_at_start => {
        "program t\ninteger :: a(6) = [2, 2, 2, 5, 6, 7]\nprint *, minval(a)\nend program t\n",
        ["2"]
    };
    maxval_int_sparse_peaks => {
        "program t\ninteger :: a(7) = [0, 0, 15, 0, 0, 20, 0]\nprint *, maxval(a)\nend program t\n",
        ["20"]
    };
    minval_int_sparse_dips => {
        "program t\ninteger :: a(7) = [100, 100, 3, 100, 100, 1, 100]\nprint *, minval(a)\nend program t\n",
        ["1"]
    };

    // ── Whole-array MINVAL/MAXVAL reals (6) ───────────────────────────

    maxval_real_positive_fractions => {
        "program t\nreal :: a(4) = [0.5, 2.5, 1.5, 3.5]\nprint *, int(maxval(a) * 10)\nend program t\n",
        ["35"]
    };
    minval_real_positive_fractions => {
        "program t\nreal :: a(4) = [0.5, 2.5, 1.5, 3.5]\nprint *, int(minval(a) * 10)\nend program t\n",
        ["5"]
    };
    maxval_real_mixed_sign => {
        "program t\nreal :: a(5) = [-2.0, 3.0, -1.0, 4.0, 0.0]\nprint *, int(maxval(a))\nend program t\n",
        ["4"]
    };
    minval_real_mixed_sign => {
        "program t\nreal :: a(5) = [-2.0, 3.0, -1.0, 4.0, 0.0]\nprint *, int(minval(a))\nend program t\n",
        ["-2"]
    };
    maxval_real_all_equal => {
        "program t\nreal :: a(3) = [2.5, 2.5, 2.5]\nprint *, int(maxval(a) * 10)\nend program t\n",
        ["25"]
    };
    minval_real_tenths => {
        "program t\nreal :: a(5) = [0.3, 0.1, 0.4, 0.2, 0.5]\nprint *, int(minval(a) * 10)\nend program t\n",
        ["1"]
    };

    // ── MAXVAL/MINVAL with DIM=1 on 2D (10) ───────────────────────────

    maxval_dim1_col1_of_3x3 => {
        "program t\ninteger :: m(3,3) = reshape([1,9,2, 8,3,7, 4,6,5], [3,3])\ninteger :: col(3)\ncol = maxval(m, dim=1)\nprint *, col(1)\nprint *, col(2)\nprint *, col(3)\nend program t\n",
        ["8", "9", "7"]
    };
    maxval_dim1_col3_of_3x3 => {
        "program t\ninteger :: m(3,3) = reshape([1,9,2, 8,3,7, 4,6,5], [3,3])\ninteger :: col(3)\ncol = maxval(m, dim=1)\nprint *, col(3)\nend program t\n",
        ["7"]
    };
    minval_dim1_col2_of_3x3 => {
        "program t\ninteger :: m(3,3) = reshape([1,9,2, 8,3,7, 4,6,5], [3,3])\ninteger :: col(3)\ncol = minval(m, dim=1)\nprint *, col(1)\nprint *, col(2)\nprint *, col(3)\nend program t\n",
        ["1", "3", "2"]
    };
    maxval_dim1_2x4_matrix => {
        "program t\ninteger :: m(2,4) = reshape([1,5,3,7, 2,6,4,8], [2,4])\ninteger :: col(4)\ncol = maxval(m, dim=1)\nprint *, col(1)\nprint *, col(4)\nend program t\n",
        ["2", "8"]
    };
    minval_dim1_2x4_matrix => {
        "program t\ninteger :: m(2,4) = reshape([1,5,3,7, 2,6,4,8], [2,4])\ninteger :: col(4)\ncol = minval(m, dim=1)\nprint *, col(2)\nprint *, col(3)\nend program t\n",
        ["5", "3"]
    };
    maxval_dim1_4x2_matrix => {
        "program t\ninteger :: m(4,2) = reshape([10,20, 30,40, 50,60, 70,80], [4,2])\ninteger :: col(2)\ncol = maxval(m, dim=1)\nprint *, col(1)\nprint *, col(2)\nend program t\n",
        ["70", "80"]
    };
    minval_dim1_4x2_matrix => {
        "program t\ninteger :: m(4,2) = reshape([10,20, 30,40, 50,60, 70,80], [4,2])\ninteger :: col(2)\ncol = minval(m, dim=1)\nprint *, col(1)\nprint *, col(2)\nend program t\n",
        ["10", "20"]
    };
    maxval_dim1_sum_of_column_maxes => {
        "program t\ninteger :: m(3,3) = reshape([(i, i=1,9)], [3,3])\ninteger :: col(3)\ncol = maxval(m, dim=1)\nprint *, sum(col)\nend program t\n",
        ["18"]
    };
    minval_dim1_sum_of_column_mins => {
        "program t\ninteger :: m(3,3) = reshape([(i, i=1,9)], [3,3])\ninteger :: col(3)\ncol = minval(m, dim=1)\nprint *, sum(col)\nend program t\n",
        ["12"]
    };
    maxval_dim1_3x2_negatives => {
        "program t\ninteger :: m(3,2) = reshape([-1,-5, -2,-3, -4,-6], [3,2])\ninteger :: col(2)\ncol = maxval(m, dim=1)\nprint *, col(1)\nprint *, col(2)\nend program t\n",
        ["-1", "-3"]
    };

    // ── MAXVAL/MINVAL with DIM=2 on 2D (8) ────────────────────────────

    maxval_dim2_row1_of_3x3 => {
        "program t\ninteger :: m(3,3) = reshape([1,9,2, 8,3,7, 4,6,5], [3,3])\ninteger :: row(3)\nrow = maxval(m, dim=2)\nprint *, row(1)\nprint *, row(2)\nprint *, row(3)\nend program t\n",
        ["9", "8", "6"]
    };
    minval_dim2_row3_of_3x3 => {
        "program t\ninteger :: m(3,3) = reshape([1,9,2, 8,3,7, 4,6,5], [3,3])\ninteger :: row(3)\nrow = minval(m, dim=2)\nprint *, row(3)\nend program t\n",
        ["4"]
    };
    maxval_dim2_2x5_matrix => {
        "program t\ninteger :: m(2,5) = reshape([1,3,5,7,9, 2,4,6,8,10], [2,5])\ninteger :: row(2)\nrow = maxval(m, dim=2)\nprint *, row(1)\nprint *, row(2)\nend program t\n",
        ["9", "10"]
    };
    minval_dim2_2x5_matrix => {
        "program t\ninteger :: m(2,5) = reshape([1,3,5,7,9, 2,4,6,8,10], [2,5])\ninteger :: row(2)\nrow = minval(m, dim=2)\nprint *, row(1)\nprint *, row(2)\nend program t\n",
        ["1", "2"]
    };
    maxval_dim2_5x2_matrix => {
        "program t\ninteger :: m(5,2) = reshape([11,12, 21,22, 31,32, 41,42, 51,52], [5,2])\ninteger :: row(5)\nrow = maxval(m, dim=2)\nprint *, row(1)\nprint *, row(5)\nend program t\n",
        ["12", "52"]
    };
    minval_dim2_5x2_matrix => {
        "program t\ninteger :: m(5,2) = reshape([11,12, 21,22, 31,32, 41,42, 51,52], [5,2])\ninteger :: row(5)\nrow = minval(m, dim=2)\nprint *, row(3)\nend program t\n",
        ["31"]
    };
    maxval_dim2_real_2x3 => {
        "program t\nreal :: m(2,3) = reshape([1.0, 3.0, 2.0, 6.0, 4.0, 5.0], [2,3])\nreal :: row(2)\nrow = maxval(m, dim=2)\nprint *, int(row(1))\nprint *, int(row(2))\nend program t\n",
        ["3", "6"]
    };
    minval_dim2_real_2x3 => {
        "program t\nreal :: m(2,3) = reshape([1.0, 3.0, 2.0, 6.0, 4.0, 5.0], [2,3])\nreal :: row(2)\nrow = minval(m, dim=2)\nprint *, int(row(1))\nprint *, int(row(2))\nend program t\n",
        ["1", "4"]
    };

    // ── MASK= on whole-array reductions (10) ──────────────────────────

    maxval_mask_skip_first_and_last => {
        "program t\ninteger :: a(7) = [100, 5, 20, 35, 50, 65, 200]\nlogical :: mask(7) = [.false., .true., .true., .true., .true., .true., .false.]\nprint *, maxval(a, mask=mask)\nend program t\n",
        ["65"]
    };
    minval_mask_skip_first_and_last => {
        "program t\ninteger :: a(7) = [100, 5, 20, 35, 50, 65, 200]\nlogical :: mask(7) = [.false., .true., .true., .true., .true., .true., .false.]\nprint *, minval(a, mask=mask)\nend program t\n",
        ["5"]
    };
    maxval_mask_positive_only => {
        "program t\ninteger :: a(6) = [-5, 3, -2, 8, -1, 4]\nlogical :: mask(6) = [.false., .true., .false., .true., .false., .true.]\nprint *, maxval(a, mask=mask)\nend program t\n",
        ["8"]
    };
    minval_mask_positive_only => {
        "program t\ninteger :: a(6) = [-5, 3, -2, 8, -1, 4]\nlogical :: mask(6) = [.false., .true., .false., .true., .false., .true.]\nprint *, minval(a, mask=mask)\nend program t\n",
        ["3"]
    };
    maxval_mask_odd_indices => {
        "program t\ninteger :: a(8) = [2, 4, 6, 8, 10, 12, 14, 16]\nlogical :: mask(8) = [.true., .false., .true., .false., .true., .false., .true., .false.]\nprint *, maxval(a, mask=mask)\nend program t\n",
        ["14"]
    };
    minval_mask_odd_indices => {
        "program t\ninteger :: a(8) = [2, 4, 6, 8, 10, 12, 14, 16]\nlogical :: mask(8) = [.true., .false., .true., .false., .true., .false., .true., .false.]\nprint *, minval(a, mask=mask)\nend program t\n",
        ["2"]
    };
    maxval_mask_single_element => {
        "program t\ninteger :: a(5) = [9, 8, 7, 6, 5]\nlogical :: mask(5) = [.false., .false., .true., .false., .false.]\nprint *, maxval(a, mask=mask)\nend program t\n",
        ["7"]
    };
    minval_mask_single_element => {
        "program t\ninteger :: a(5) = [9, 8, 7, 6, 5]\nlogical :: mask(5) = [.false., .false., .true., .false., .false.]\nprint *, minval(a, mask=mask)\nend program t\n",
        ["7"]
    };
    maxval_mask_real_values => {
        "program t\nreal :: a(5) = [1.1, 2.2, 3.3, 4.4, 5.5]\nlogical :: mask(5) = [.true., .false., .true., .false., .true.]\nprint *, int(maxval(a, mask=mask) * 10)\nend program t\n",
        ["55"]
    };
    minval_mask_real_values => {
        "program t\nreal :: a(5) = [1.1, 2.2, 3.3, 4.4, 5.5]\nlogical :: mask(5) = [.true., .false., .true., .false., .true.]\nprint *, int(minval(a, mask=mask) * 10)\nend program t\n",
        ["11"]
    };

    // ── MAXLOC/MINLOC whole array known positions (10) ────────────────

    maxloc_int_peak_at_index_four => {
        "program t\ninteger :: a(6) = [2, 5, 3, 9, 1, 7]\ninteger :: loc(1)\nloc = maxloc(a)\nprint *, loc(1)\nprint *, a(loc(1))\nend program t\n",
        ["4", "9"]
    };
    minloc_int_dip_at_index_two => {
        "program t\ninteger :: a(6) = [8, 1, 6, 4, 9, 3]\ninteger :: loc(1)\nloc = minloc(a)\nprint *, loc(1)\nprint *, a(loc(1))\nend program t\n",
        ["2", "1"]
    };
    maxloc_int_tie_first_wins_at_two => {
        "program t\ninteger :: a(5) = [4, 7, 7, 3, 7]\ninteger :: loc(1)\nloc = maxloc(a)\nprint *, loc(1)\nend program t\n",
        ["2"]
    };
    minloc_int_tie_first_wins_at_one => {
        "program t\ninteger :: a(5) = [2, 2, 5, 2, 8]\ninteger :: loc(1)\nloc = minloc(a)\nprint *, loc(1)\nend program t\n",
        ["1"]
    };
    maxloc_real_peak_index => {
        "program t\nreal :: a(4) = [1.0, 4.0, 2.5, 3.0]\ninteger :: loc(1)\nloc = maxloc(a)\nprint *, loc(1)\nend program t\n",
        ["2"]
    };
    minloc_real_dip_index => {
        "program t\nreal :: a(4) = [1.0, 4.0, 2.5, 3.0]\ninteger :: loc(1)\nloc = minloc(a)\nprint *, loc(1)\nend program t\n",
        ["1"]
    };
    maxloc_mask_excludes_global_max => {
        "program t\ninteger :: a(5) = [10, 20, 30, 40, 50]\nlogical :: mask(5) = [.true., .true., .true., .false., .false.]\ninteger :: loc(1)\nloc = maxloc(a, mask=mask)\nprint *, loc(1)\nprint *, a(loc(1))\nend program t\n",
        ["3", "30"]
    };
    minloc_mask_excludes_global_min => {
        "program t\ninteger :: a(5) = [10, 20, 30, 40, 50]\nlogical :: mask(5) = [.false., .true., .true., .true., .true.]\ninteger :: loc(1)\nloc = minloc(a, mask=mask)\nprint *, loc(1)\nprint *, a(loc(1))\nend program t\n",
        ["2", "20"]
    };
    maxloc_1d_last_element => {
        "program t\ninteger :: a(4) = [1, 2, 3, 99]\ninteger :: loc(1)\nloc = maxloc(a)\nprint *, loc(1)\nend program t\n",
        ["4"]
    };
    minloc_1d_last_element => {
        "program t\ninteger :: a(4) = [9, 8, 7, 1]\ninteger :: loc(1)\nloc = minloc(a)\nprint *, loc(1)\nend program t\n",
        ["4"]
    };

    // ── MAXLOC/MINLOC with DIM= (8) ───────────────────────────────────

    maxloc_dim1_second_column => {
        "program t\ninteger :: m(3,3) = reshape([1,9,2, 8,3,7, 4,6,5], [3,3])\ninteger :: col(3)\ncol = maxloc(m, dim=1)\nprint *, col(2)\nend program t\n",
        ["1"]
    };
    minloc_dim1_third_column => {
        "program t\ninteger :: m(3,3) = reshape([1,9,2, 8,3,7, 4,6,5], [3,3])\ninteger :: col(3)\ncol = minloc(m, dim=1)\nprint *, col(3)\nend program t\n",
        ["1"]
    };
    maxloc_dim2_middle_row => {
        "program t\ninteger :: m(3,3) = reshape([1,9,2, 8,3,7, 4,6,5], [3,3])\ninteger :: row(3)\nrow = maxloc(m, dim=2)\nprint *, row(2)\nend program t\n",
        ["1"]
    };
    minloc_dim2_bottom_row => {
        "program t\ninteger :: m(3,3) = reshape([1,9,2, 8,3,7, 4,6,5], [3,3])\ninteger :: row(3)\nrow = minloc(m, dim=2)\nprint *, row(3)\nend program t\n",
        ["1"]
    };
    maxloc_dim1_with_mask_column => {
        "program t\ninteger :: m(2,3) = reshape([1, 100, 3, 4, 5, 6], [2,3])\nlogical :: mask(2,3) = reshape([.true., .false., .true., .true., .true., .true.], [2,3])\ninteger :: col(3)\ncol = maxloc(m, dim=1, mask=mask)\nprint *, col(1)\nprint *, col(2)\nend program t\n",
        ["2", "2"]
    };
    minloc_dim1_with_mask_column => {
        "program t\ninteger :: m(2,3) = reshape([1, 100, 3, 4, 5, 6], [2,3])\nlogical :: mask(2,3) = reshape([.true., .false., .true., .true., .true., .true.], [2,3])\ninteger :: col(3)\ncol = minloc(m, dim=1, mask=mask)\nprint *, col(1)\nprint *, col(3)\nend program t\n",
        ["1", "1"]
    };
    maxloc_dim2_with_mask_row => {
        "program t\ninteger :: m(3,2) = reshape([10, 1, 20, 2, 30, 3], [3,2])\nlogical :: mask(3,2) = reshape([.true., .false., .true., .false., .true., .false.], [3,2])\ninteger :: row(3)\nrow = maxloc(m, dim=2, mask=mask)\nprint *, row(1)\nprint *, row(3)\nend program t\n",
        ["1", "1"]
    };
    minloc_dim2_with_mask_row => {
        "program t\ninteger :: m(3,2) = reshape([10, 1, 20, 2, 30, 3], [3,2])\nlogical :: mask(3,2) = reshape([.true., .false., .true., .false., .true., .false.], [3,2])\ninteger :: row(3)\nrow = minloc(m, dim=2, mask=mask)\nprint *, row(2)\nend program t\n",
        ["1"]
    };
}
