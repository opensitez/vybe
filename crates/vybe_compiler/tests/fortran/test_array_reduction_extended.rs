//! Extended array reduction coverage: sum, product, count, any, and all with dim=,
//! mask=, slices, logical arrays, and real arrays with known totals.
//! Distinct from `test_arrays.rs` (basic whole-array reductions) and
//! `test_arrays_dim_mask.rs` (dim/mask compile-only smoke tests).

fortran_cases! {
    // ── Integer sum with known totals (5) ───────────────────────────

    sum_int_one_to_six => {
        "program t\ninteger :: a(6) = [(i, i = 1, 6)]\nprint *, sum(a)\nend program t\n",
        ["21"]
    };
    sum_int_alternating_signs => {
        "program t\ninteger :: a(6) = [5, -2, 8, -1, 3, -4]\nprint *, sum(a)\nend program t\n",
        ["9"]
    };
    sum_int_constant_vector => {
        "program t\ninteger :: a(5) = [7, 7, 7, 7, 7]\nprint *, sum(a)\nend program t\n",
        ["35"]
    };
    sum_int_sparse_nonzeros => {
        "program t\ninteger :: a(8) = [0, 0, 10, 0, 20, 0, 30, 0]\nprint *, sum(a)\nend program t\n",
        ["60"]
    };
    sum_int_triangular_ten => {
        "program t\ninteger :: a(4) = [1, 3, 6, 10]\nprint *, sum(a)\nend program t\n",
        ["20"]
    };

    // ── Real sum with known totals (5) ──────────────────────────────

    sum_real_unit_interval => {
        "program t\nreal :: a(4) = [0.5, 1.5, 2.5, 3.5]\nprint *, sum(a)\nend program t\n",
        ["8"]
    };
    sum_real_tenths_one_to_five => {
        "program t\nreal :: a(5) = [0.1, 0.2, 0.3, 0.4, 0.5]\nprint *, sum(a)\nend program t\n",
        ["1.5"]
    };
    sum_real_three_halves => {
        "program t\nreal :: a(3) = [1.5, 2.5, 3.5]\nprint *, sum(a)\nend program t\n",
        ["7.5"]
    };
    sum_real_negative_mix => {
        "program t\nreal :: a(4) = [2.0, -1.0, 3.0, -2.0]\nprint *, sum(a)\nend program t\n",
        ["2"]
    };
    sum_real_six_values_twenty_one => {
        "program t\nreal :: a(6) = [1.0, 2.0, 3.0, 4.0, 5.0, 6.0]\nprint *, sum(a)\nend program t\n",
        ["21"]
    };

    // ── Product reductions (5) ──────────────────────────────────────

    product_int_one_to_five => {
        "program t\ninteger :: a(5) = [(i, i = 1, 5)]\nprint *, product(a)\nend program t\n",
        ["120"]
    };
    product_int_powers_of_two => {
        "program t\ninteger :: a(4) = [2, 4, 8, 16]\nprint *, product(a)\nend program t\n",
        ["1024"]
    };
    product_int_with_zero => {
        "program t\ninteger :: a(5) = [3, 5, 0, 7, 9]\nprint *, product(a)\nend program t\n",
        ["0"]
    };
    product_real_three_halves => {
        "program t\nreal :: a(3) = [2.0, 2.5, 2.0]\nprint *, product(a)\nend program t\n",
        ["10"]
    };
    product_int_negatives_pair => {
        "program t\ninteger :: a(4) = [-2, 3, -4, 5]\nprint *, product(a)\nend program t\n",
        ["120"]
    };

    // ── COUNT on logical and comparison masks (5) ───────────────────

    count_logical_three_of_five => {
        "program t\nlogical :: m(5) = [.true., .false., .true., .false., .true.]\nprint *, count(m)\nend program t\n",
        ["3"]
    };
    count_logical_all_false => {
        "program t\nlogical :: m(4) = [.false., .false., .false., .false.]\nprint *, count(m)\nend program t\n",
        ["0"]
    };
    count_int_greater_than_four => {
        "program t\ninteger :: a(7) = [1, 5, 3, 8, 2, 6, 4]\nprint *, count(a > 4)\nend program t\n",
        ["3"]
    };
    count_int_equal_to_two => {
        "program t\ninteger :: a(6) = [1, 2, 2, 3, 2, 4]\nprint *, count(a == 2)\nend program t\n",
        ["3"]
    };
    count_real_positive_values => {
        "program t\nreal :: a(5) = [-1.0, 0.0, 1.5, -2.0, 3.0]\nprint *, count(a > 0.0)\nend program t\n",
        ["2"]
    };

    // ── ANY / ALL on logical arrays (5) ─────────────────────────────

    any_logical_middle_true => {
        "program t\nlogical :: m(5) = [.false., .false., .true., .false., .false.]\nprint *, any(m)\nend program t\n",
        ["true"]
    };
    any_logical_all_false => {
        "program t\nlogical :: m(3) = [.false., .false., .false.]\nprint *, any(m)\nend program t\n",
        ["false"]
    };
    all_logical_all_true => {
        "program t\nlogical :: m(4) = [.true., .true., .true., .true.]\nprint *, all(m)\nend program t\n",
        ["true"]
    };
    all_logical_one_false => {
        "program t\nlogical :: m(4) = [.true., .true., .false., .true.]\nprint *, all(m)\nend program t\n",
        ["false"]
    };
    any_all_on_comparison_array => {
        "program t\ninteger :: a(5) = [2, 4, 6, 8, 10]\nprint *, any(a > 7)\nprint *, all(a > 0)\nend program t\n",
        ["true", "true"]
    };

    // ── SUM with dim= (5) ───────────────────────────────────────────

    sum_dim1_two_by_three_cols => {
        "program t\ninteger :: m(2,3) = reshape([1,2,3,4,5,6],[2,3])\ninteger :: c(3)\nc = sum(m, dim=1)\nprint *, c(1)\nprint *, c(2)\nprint *, c(3)\nend program t\n",
        ["5", "7", "9"]
    };
    sum_dim2_two_by_three_rows => {
        "program t\ninteger :: m(2,3) = reshape([1,2,3,4,5,6],[2,3])\ninteger :: r(2)\nr = sum(m, dim=2)\nprint *, r(1)\nprint *, r(2)\nend program t\n",
        ["6", "15"]
    };
    sum_dim1_three_by_two => {
        "program t\ninteger :: m(3,2) = reshape([1,2,3,4,5,6],[3,2])\ninteger :: c(2)\nc = sum(m, dim=1)\nprint *, c(1)\nprint *, c(2)\nend program t\n",
        ["9", "12"]
    };
    sum_dim2_three_by_four_rows => {
        "program t\ninteger :: m(3,4) = reshape([(i, i = 1, 12)],[3,4])\ninteger :: r(3)\nr = sum(m, dim=2)\nprint *, r(1)\nprint *, r(2)\nprint *, r(3)\nend program t\n",
        ["10", "26", "42"]
    };
    sum_real_dim1_two_by_four => {
        "program t\nreal :: m(2,4) = reshape([1.0,2.0,3.0,4.0,5.0,6.0,7.0,8.0],[2,4])\nreal :: c(4)\nc = sum(m, dim=1)\nprint *, c(1)\nprint *, c(4)\nend program t\n",
        ["6", "12"]
    };

    // ── PRODUCT with dim= (3) ───────────────────────────────────────

    product_dim1_two_by_three => {
        "program t\ninteger :: m(2,3) = reshape([1,2,3,4,5,6],[2,3])\ninteger :: c(3)\nc = product(m, dim=1)\nprint *, c(1)\nprint *, c(2)\nprint *, c(3)\nend program t\n",
        ["4", "10", "18"]
    };
    product_dim2_two_by_three => {
        "program t\ninteger :: m(2,3) = reshape([1,2,3,4,5,6],[2,3])\ninteger :: r(2)\nr = product(m, dim=2)\nprint *, r(1)\nprint *, r(2)\nend program t\n",
        ["6", "120"]
    };
    product_dim2_three_by_two => {
        "program t\ninteger :: m(3,2) = reshape([2,3,4,5,6,7],[3,2])\ninteger :: r(3)\nr = product(m, dim=2)\nprint *, r(1)\nprint *, r(2)\nprint *, r(3)\nend program t\n",
        ["6", "20", "42"]
    };

    // ── COUNT with dim= (3) ─────────────────────────────────────────

    count_dim1_int_gt_six => {
        "program t\ninteger :: m(3,4) = reshape([(i, i = 1, 12)],[3,4])\ninteger :: c(4)\nc = count(m > 6, dim=1)\nprint *, c(1)\nprint *, c(3)\nprint *, c(4)\nend program t\n",
        ["1", "2", "2"]
    };
    count_dim2_logical_rows => {
        "program t\nlogical :: m(2,3) = reshape([.true.,.false.,.true.,.false.,.true.,.false.],[2,3])\ninteger :: r(2)\nr = count(m, dim=2)\nprint *, r(1)\nprint *, r(2)\nend program t\n",
        ["2", "1"]
    };
    count_dim1_real_positive => {
        "program t\nreal :: m(2,3) = reshape([1.0,-1.0,2.0,3.0,-2.0,4.0],[2,3])\ninteger :: c(3)\nc = count(m > 0.0, dim=1)\nprint *, c(1)\nprint *, c(2)\nprint *, c(3)\nend program t\n",
        ["2", "1", "2"]
    };

    // ── ANY / ALL with dim= (4) ─────────────────────────────────────

    any_dim1_logical_columns => {
        "program t\nlogical :: m(2,3) = reshape([.false.,.false.,.true.,.false.,.false.,.false.],[2,3])\nlogical :: c(3)\nc = any(m, dim=1)\nprint *, c(1)\nprint *, c(2)\nprint *, c(3)\nend program t\n",
        ["false", "false", "true"]
    };
    any_dim2_logical_rows => {
        "program t\nlogical :: m(3,2) = reshape([.false.,.true.,.false.,.false.,.false.,.false.],[3,2])\nlogical :: r(3)\nr = any(m, dim=2)\nprint *, r(1)\nprint *, r(2)\nprint *, r(3)\nend program t\n",
        ["true", "false", "false"]
    };
    all_dim1_logical_columns => {
        "program t\nlogical :: m(2,3) = reshape([.true.,.true.,.true.,.false.,.true.,.true.],[2,3])\nlogical :: c(3)\nc = all(m, dim=1)\nprint *, c(1)\nprint *, c(2)\nprint *, c(3)\nend program t\n",
        ["false", "true", "true"]
    };
    all_dim2_logical_rows => {
        "program t\nlogical :: m(3,2) = reshape([.true.,.true.,.true.,.false.,.true.,.false.],[3,2])\nlogical :: r(3)\nr = all(m, dim=2)\nprint *, r(1)\nprint *, r(2)\nprint *, r(3)\nend program t\n",
        ["true", "false", "false"]
    };

    // ── mask= on whole arrays (5) ───────────────────────────────────

    sum_mask_even_positions => {
        "program t\ninteger :: a(6) = [1, 2, 3, 4, 5, 6]\nlogical :: m(6) = [.false., .true., .false., .true., .false., .true.]\nprint *, sum(a, mask=m)\nend program t\n",
        ["12"]
    };
    product_mask_select_three => {
        "program t\ninteger :: a(5) = [1, 2, 3, 4, 5]\nlogical :: m(5) = [.true., .true., .false., .true., .false.]\nprint *, product(a, mask=m)\nend program t\n",
        ["8"]
    };
    count_explicit_mask_array => {
        "program t\nlogical :: m(6) = [.true., .false., .true., .true., .false., .false.]\nprint *, count(m)\nend program t\n",
        ["3"]
    };
    any_mask_last_two => {
        "program t\nlogical :: m(4) = [.false., .false., .true., .true.]\nprint *, any(m)\nend program t\n",
        ["true"]
    };
    all_mask_first_three => {
        "program t\nlogical :: m(5) = [.true., .true., .true., .false., .true.]\nprint *, all(m)\nend program t\n",
        ["false"]
    };

    // ── dim= combined with mask= (5) ────────────────────────────────

    sum_dim1_mask_two_by_four => {
        "program t\ninteger :: m(2,4) = reshape([1,2,3,4,5,6,7,8],[2,4])\nlogical :: mask(2,4) = reshape([.true.,.false.,.true.,.false.,.true.,.false.,.true.,.false.],[2,4])\ninteger :: c(4)\nc = sum(m, dim=1, mask=mask)\nprint *, c(1)\nprint *, c(3)\nend program t\n",
        ["6", "10"]
    };
    sum_dim2_mask_three_by_three => {
        "program t\ninteger :: m(3,3) = reshape([(i, i = 1, 9)],[3,3])\nlogical :: mask(3,3)\nmask = m > 4\ninteger :: r(3)\nr = sum(m, dim=2, mask=mask)\nprint *, r(1)\nprint *, r(3)\nend program t\n",
        ["0", "30"]
    };
    product_dim2_mask_two_by_three => {
        "program t\ninteger :: m(2,3) = reshape([2,3,4,5,6,7],[2,3])\nlogical :: mask(2,3) = reshape([.true.,.false.,.true.,.false.,.true.,.false.],[2,3])\ninteger :: r(2)\nr = product(m, dim=2, mask=mask)\nprint *, r(1)\nprint *, r(2)\nend program t\n",
        ["8", "30"]
    };
    count_dim1_mask_comparison => {
        "program t\ninteger :: m(2,4) = reshape([1,5,2,8,3,6,4,7],[2,4])\nlogical :: mask(2,4) = reshape([.true.,.true.,.false.,.true.,.false.,.true.,.true.,.false.],[2,4])\ninteger :: c(4)\nc = count(m > 4, dim=1, mask=mask)\nprint *, c(1)\nprint *, c(4)\nend program t\n",
        ["1", "1"]
    };
    sum_real_dim1_mask_fractions => {
        "program t\nreal :: m(2,3) = reshape([1.0,2.0,3.0,4.0,5.0,6.0],[2,3])\nlogical :: mask(2,3) = reshape([.true.,.false.,.true.,.false.,.true.,.false.],[2,3])\nreal :: c(3)\nc = sum(m, dim=1, mask=mask)\nprint *, c(1)\nprint *, c(3)\nend program t\n",
        ["5", "9"]
    };

    // ── Reductions on array slices (5) ──────────────────────────────

    sum_slice_three_to_seven => {
        "program t\ninteger :: a(9) = [(i * 2, i = 1, 9)]\nprint *, sum(a(3:7))\nend program t\n",
        ["50"]
    };
    product_slice_two_to_five => {
        "program t\ninteger :: a(8) = [(i, i = 1, 8)]\nprint *, product(a(3:5))\nend program t\n",
        ["60"]
    };
    count_gt_on_slice => {
        "program t\ninteger :: a(7) = [1, 2, 3, 4, 5, 6, 7]\nprint *, count(a(2:6) > 3)\nend program t\n",
        ["3"]
    };
    any_logical_slice_middle => {
        "program t\nlogical :: m(5) = [.true., .false., .true., .false., .true.]\nprint *, any(m(2:4))\nend program t\n",
        ["true"]
    };
    all_logical_slice_all_true => {
        "program t\nlogical :: m(6) = [.true., .true., .true., .true., .false., .true.]\nprint *, all(m(1:4))\nend program t\n",
        ["true"]
    };
}
