//! Fortran array locators (maxloc, minloc, findloc) and masked reductions.
//! Distinct from `test_arrays.rs` bare maxval/minval and `test_arrays_dim_mask.rs` compile-only dim/mask cases.

fortran_cases! {
    maxloc_1d_unique_peak_at_index_three => {
        "program t\ninteger :: a(5) = [3, 1, 9, 1, 5]\ninteger :: loc(1)\nloc = maxloc(a)\nprint *, loc(1)\nend program t\n",
        ["3"]
    };

    maxloc_1d_tie_returns_first_occurrence => {
        "program t\ninteger :: a(5) = [5, 3, 5, 2, 5]\ninteger :: loc(1)\nloc = maxloc(a)\nprint *, loc(1)\nend program t\n",
        ["1"]
    };

    maxloc_1d_negative_values_peak_at_start => {
        "program t\ninteger :: a(5) = [-1, -5, -2, -5, -3]\ninteger :: loc(1)\nloc = maxloc(a)\nprint *, loc(1)\nend program t\n",
        ["1"]
    };

    maxloc_1d_all_elements_equal_returns_one => {
        "program t\ninteger :: a(3) = [7, 7, 7]\ninteger :: loc(1)\nloc = maxloc(a)\nprint *, loc(1)\nend program t\n",
        ["1"]
    };

    maxloc_1d_single_element_array => {
        "program t\ninteger :: a(1) = [42]\ninteger :: loc(1)\nloc = maxloc(a)\nprint *, loc(1)\nend program t\n",
        ["1"]
    };

    maxloc_1d_maximum_at_first_position => {
        "program t\ninteger :: a(4) = [99, 1, 2, 3]\ninteger :: loc(1)\nloc = maxloc(a)\nprint *, loc(1)\nend program t\n",
        ["1"]
    };

    maxloc_1d_maximum_at_last_position => {
        "program t\ninteger :: a(4) = [1, 2, 3, 99]\ninteger :: loc(1)\nloc = maxloc(a)\nprint *, loc(1)\nend program t\n",
        ["4"]
    };

    maxloc_1d_descending_first_is_max => {
        "program t\ninteger :: a(5) = [10, 8, 6, 4, 2]\ninteger :: loc(1)\nloc = maxloc(a)\nprint *, loc(1)\nend program t\n",
        ["1"]
    };

    maxloc_1d_sparse_peak_at_four => {
        "program t\ninteger :: a(5) = [0, 0, 0, 5, 0]\ninteger :: loc(1)\nloc = maxloc(a)\nprint *, loc(1)\nend program t\n",
        ["4"]
    };

    maxloc_1d_plateau_at_start => {
        "program t\ninteger :: a(6) = [4, 4, 4, 4, 4, 4]\ninteger :: loc(1)\nloc = maxloc(a)\nprint *, loc(1)\nend program t\n",
        ["1"]
    };

    maxloc_1d_unimodal_peak_at_four => {
        "program t\ninteger :: a(7) = [1, 3, 5, 7, 6, 4, 2]\ninteger :: loc(1)\nloc = maxloc(a)\nprint *, loc(1)\nend program t\n",
        ["4"]
    };

    maxloc_1d_trailing_nine_at_five => {
        "program t\ninteger :: a(5) = [8, 2, 6, 2, 9]\ninteger :: loc(1)\nloc = maxloc(a)\nprint *, loc(1)\nend program t\n",
        ["5"]
    };

    minloc_1d_unique_minimum_at_two => {
        "program t\ninteger :: a(5) = [3, 1, 9, 1, 5]\ninteger :: loc(1)\nloc = minloc(a)\nprint *, loc(1)\nend program t\n",
        ["2"]
    };

    minloc_1d_tie_returns_first_occurrence => {
        "program t\ninteger :: a(5) = [5, 1, 3, 1, 4]\ninteger :: loc(1)\nloc = minloc(a)\nprint *, loc(1)\nend program t\n",
        ["2"]
    };

    minloc_1d_negative_minimum_at_two => {
        "program t\ninteger :: a(3) = [-1, -5, -2]\ninteger :: loc(1)\nloc = minloc(a)\nprint *, loc(1)\nend program t\n",
        ["2"]
    };

    minloc_1d_all_equal_returns_one => {
        "program t\ninteger :: a(3) = [4, 4, 4]\ninteger :: loc(1)\nloc = minloc(a)\nprint *, loc(1)\nend program t\n",
        ["1"]
    };

    minloc_1d_minimum_at_last_index => {
        "program t\ninteger :: a(5) = [5, 4, 3, 2, 1]\ninteger :: loc(1)\nloc = minloc(a)\nprint *, loc(1)\nend program t\n",
        ["5"]
    };

    minloc_1d_minimum_at_first_index => {
        "program t\ninteger :: a(5) = [1, 5, 4, 3, 2]\ninteger :: loc(1)\nloc = minloc(a)\nprint *, loc(1)\nend program t\n",
        ["1"]
    };

    minloc_1d_ascending_first_is_min => {
        "program t\ninteger :: a(5) = [1, 2, 3, 4, 5]\ninteger :: loc(1)\nloc = minloc(a)\nprint *, loc(1)\nend program t\n",
        ["1"]
    };

    minloc_1d_plateau_at_start => {
        "program t\ninteger :: a(6) = [4, 4, 4, 4, 4, 4]\ninteger :: loc(1)\nloc = minloc(a)\nprint *, loc(1)\nend program t\n",
        ["1"]
    };

    minloc_1d_unimodal_minimum_at_one => {
        "program t\ninteger :: a(7) = [1, 3, 5, 7, 6, 4, 2]\ninteger :: loc(1)\nloc = minloc(a)\nprint *, loc(1)\nend program t\n",
        ["1"]
    };

    minloc_1d_leading_two_at_index_two => {
        "program t\ninteger :: a(5) = [8, 2, 6, 2, 9]\ninteger :: loc(1)\nloc = minloc(a)\nprint *, loc(1)\nend program t\n",
        ["2"]
    };

    maxloc_prints_index_and_value_together => {
        "program t\ninteger :: a(5) = [3, 1, 9, 1, 5]\ninteger :: loc(1)\nloc = maxloc(a)\nprint *, loc(1), a(loc(1))\nend program t\n",
        ["3 9"]
    };

    minloc_prints_index_and_value_together => {
        "program t\ninteger :: a(5) = [3, 1, 9, 1, 5]\ninteger :: loc(1)\nloc = minloc(a)\nprint *, loc(1), a(loc(1))\nend program t\n",
        ["2 1"]
    };

    maxloc_then_value_on_separate_lines => {
        "program t\ninteger :: a(5) = [8, 2, 6, 2, 9]\ninteger :: loc(1)\nloc = maxloc(a)\nprint *, loc(1)\nprint *, a(loc(1))\nend program t\n",
        ["5", "9"]
    };

    minloc_then_value_on_separate_lines => {
        "program t\ninteger :: a(5) = [8, 2, 6, 2, 9]\ninteger :: loc(1)\nloc = minloc(a)\nprint *, loc(1)\nprint *, a(loc(1))\nend program t\n",
        ["2", "2"]
    };

    maxloc_value_at_seven_array => {
        "program t\ninteger :: a(5) = [2, 7, 1, 7, 3]\ninteger :: loc(1)\nloc = maxloc(a)\nprint *, a(loc(1))\nend program t\n",
        ["7"]
    };

    minloc_value_at_one_array => {
        "program t\ninteger :: a(5) = [2, 7, 1, 7, 3]\ninteger :: loc(1)\nloc = minloc(a)\nprint *, a(loc(1))\nend program t\n",
        ["1"]
    };

    maxval_masked_skips_leading_nine => {
        "program t\ninteger :: a(6) = [1, 9, 2, 8, 3, 7]\nlogical :: mask(6) = [.false., .false., .true., .true., .true., .true.]\nprint *, maxval(a, mask=mask)\nend program t\n",
        ["9"]
    };

    minval_masked_skips_small_unmasked => {
        "program t\ninteger :: a(6) = [10, 1, 20, 2, 30, 3]\nlogical :: mask(6) = [.true., .false., .true., .false., .true., .false.]\nprint *, minval(a, mask=mask)\nend program t\n",
        ["1"]
    };

    maxval_masked_upper_half_only => {
        "program t\ninteger :: a(8) = [2, 4, 6, 8, 10, 12, 14, 16]\nlogical :: mask(8) = [.false., .false., .false., .false., .true., .true., .true., .true.]\nprint *, maxval(a, mask=mask)\nend program t\n",
        ["16"]
    };

    minval_masked_upper_half_only => {
        "program t\ninteger :: a(8) = [2, 4, 6, 8, 10, 12, 14, 16]\nlogical :: mask(8) = [.false., .false., .false., .false., .true., .true., .true., .true.]\nprint *, minval(a, mask=mask)\nend program t\n",
        ["10"]
    };

    maxval_masked_even_positions => {
        "program t\ninteger :: a(6) = [11, 22, 33, 44, 55, 66]\nlogical :: mask(6) = [.false., .true., .false., .true., .false., .true.]\nprint *, maxval(a, mask=mask)\nend program t\n",
        ["66"]
    };

    minval_masked_even_positions => {
        "program t\ninteger :: a(6) = [11, 22, 33, 44, 55, 66]\nlogical :: mask(6) = [.false., .true., .false., .true., .false., .true.]\nprint *, minval(a, mask=mask)\nend program t\n",
        ["22"]
    };

    maxval_masked_interior_window => {
        "program t\ninteger :: a(7) = [100, 5, 20, 35, 50, 65, 200]\nlogical :: mask(7) = [.false., .true., .true., .true., .true., .true., .false.]\nprint *, maxval(a, mask=mask)\nend program t\n",
        ["65"]
    };

    minval_masked_interior_window => {
        "program t\ninteger :: a(7) = [100, 5, 20, 35, 50, 65, 200]\nlogical :: mask(7) = [.false., .true., .true., .true., .true., .true., .false.]\nprint *, minval(a, mask=mask)\nend program t\n",
        ["5"]
    };

    findloc_mask_all_true_finds_first => {
        "program t\ninteger :: a(4) = [2, 3, 4, 5]\nlogical :: mask(4) = [.true., .true., .true., .true.]\ninteger :: loc(1)\nloc = findloc(a, 3, mask=mask)\nprint *, loc(1)\nend program t\n",
        ["2"]
    };

    findloc_mask_skips_leading_positions => {
        "program t\ninteger :: a(5) = [9, 9, 9, 4, 5]\nlogical :: mask(5) = [.false., .false., .false., .true., .true.]\ninteger :: loc(1)\nloc = findloc(a, 4, mask=mask)\nprint *, loc(1)\nend program t\n",
        ["4"]
    };

    findloc_mask_only_last_true_element => {
        "program t\ninteger :: a(5) = [1, 1, 1, 1, 9]\nlogical :: mask(5) = [.false., .false., .false., .false., .true.]\ninteger :: loc(1)\nloc = findloc(a, 9, mask=mask)\nprint *, loc(1)\nend program t\n",
        ["5"]
    };

    findloc_mask_middle_true_finds_center => {
        "program t\ninteger :: a(5) = [1, 2, 3, 2, 1]\nlogical :: mask(5) = [.false., .false., .true., .false., .false.]\ninteger :: loc(1)\nloc = findloc(a, 3, mask=mask)\nprint *, loc(1)\nend program t\n",
        ["3"]
    };

    findloc_mask_no_true_returns_zero => {
        "program t\ninteger :: a(3) = [1, 2, 3]\nlogical :: mask(3) = [.false., .false., .false.]\ninteger :: loc(1)\nloc = findloc(a, 2, mask=mask)\nprint *, loc(1)\nend program t\n",
        ["0"]
    };

    findloc_forward_first_occurrence => {
        "program t\ninteger :: a(6) = [1, 2, 1, 2, 1, 2]\ninteger :: loc(1)\nloc = findloc(a, 1)\nprint *, loc(1)\nend program t\n",
        ["1"]
    };

    findloc_back_last_occurrence => {
        "program t\ninteger :: a(6) = [1, 2, 1, 2, 1, 2]\ninteger :: loc(1)\nloc = findloc(a, 1, back=.true.)\nprint *, loc(1)\nend program t\n",
        ["5"]
    };

    findloc_back_value_two_last_index => {
        "program t\ninteger :: a(6) = [1, 2, 1, 2, 1, 2]\ninteger :: loc(1)\nloc = findloc(a, 2, back=.true.)\nprint *, loc(1)\nend program t\n",
        ["6"]
    };

    findloc_forward_value_two_first_index => {
        "program t\ninteger :: a(6) = [1, 2, 1, 2, 1, 2]\ninteger :: loc(1)\nloc = findloc(a, 2)\nprint *, loc(1)\nend program t\n",
        ["2"]
    };

    findloc_back_with_mask_restricted_tail => {
        "program t\ninteger :: a(6) = [1, 2, 1, 2, 1, 2]\nlogical :: mask(6) = [.true., .true., .true., .true., .true., .false.]\ninteger :: loc(1)\nloc = findloc(a, 1, back=.true., mask=mask)\nprint *, loc(1)\nend program t\n",
        ["5"]
    };

    maxloc_dim1_first_column_row_index => {
        "program t\ninteger :: m(3,3) = reshape([1,9,2,8,3,7,4,6,5],[3,3])\ninteger :: col_maxloc(3)\ncol_maxloc = maxloc(m, dim=1)\nprint *, col_maxloc(1)\nend program t\n",
        ["2"]
    };

    maxloc_dim1_second_column_row_index => {
        "program t\ninteger :: m(3,3) = reshape([1,9,2,8,3,7,4,6,5],[3,3])\ninteger :: col_maxloc(3)\ncol_maxloc = maxloc(m, dim=1)\nprint *, col_maxloc(2)\nend program t\n",
        ["1"]
    };

    maxloc_dim1_third_column_row_index => {
        "program t\ninteger :: m(3,3) = reshape([1,9,2,8,3,7,4,6,5],[3,3])\ninteger :: col_maxloc(3)\ncol_maxloc = maxloc(m, dim=1)\nprint *, col_maxloc(3)\nend program t\n",
        ["2"]
    };

    maxloc_dim2_first_row_column_index => {
        "program t\ninteger :: m(3,3) = reshape([1,9,2,8,3,7,4,6,5],[3,3])\ninteger :: row_maxloc(3)\nrow_maxloc = maxloc(m, dim=2)\nprint *, row_maxloc(1)\nend program t\n",
        ["2"]
    };

    minloc_dim2_first_row_column_index => {
        "program t\ninteger :: m(3,3) = reshape([1,9,2,8,3,7,4,6,5],[3,3])\ninteger :: row_minloc(3)\nrow_minloc = minloc(m, dim=2)\nprint *, row_minloc(1)\nend program t\n",
        ["1"]
    };

    where_style_mask_maxloc_positive_only => {
        "program t\ninteger :: a(5) = [3, -1, 5, -2, 4]\nlogical :: mask(5) = [.true., .false., .true., .false., .true.]\ninteger :: loc(1)\nloc = maxloc(a, mask=mask)\nprint *, loc(1)\nend program t\n",
        ["3"]
    };

    where_style_mask_minloc_positive_only => {
        "program t\ninteger :: a(5) = [3, -1, 5, -2, 4]\nlogical :: mask(5) = [.true., .false., .true., .false., .true.]\ninteger :: loc(1)\nloc = minloc(a, mask=mask)\nprint *, loc(1)\nend program t\n",
        ["1"]
    };

    where_style_mask_maxval_positive_only => {
        "program t\ninteger :: a(5) = [3, -1, 5, -2, 4]\nlogical :: mask(5) = [.true., .false., .true., .false., .true.]\nprint *, maxval(a, mask=mask)\nend program t\n",
        ["5"]
    };

    where_style_mask_findloc_value_five => {
        "program t\ninteger :: a(5) = [3, -1, 5, -2, 4]\nlogical :: mask(5) = [.true., .false., .true., .false., .true.]\ninteger :: loc(1)\nloc = findloc(a, 5, mask=mask)\nprint *, loc(1)\nend program t\n",
        ["3"]
    };
}
