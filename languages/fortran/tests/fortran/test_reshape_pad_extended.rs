//! Extended RESHAPE with ORDER=, PAD=, multidimensional totals, and size
//! mismatches. Distinct from `test_array_transforms.rs` basic 2x2 reshape and
//! `test_arrays.rs` compile-only reshape smoke.

use super::helpers::compile_ok;

fortran_cases! {
    // ── Default Fortran column-major ORDER (10) ───────────────────────

    reshape_fortran_2x3_first_column => {
        "program t\ninteger :: a(6) = [1, 2, 3, 4, 5, 6]\ninteger :: m(2,3)\nm = reshape(a, [2, 3])\nprint *, m(1,1)\nprint *, m(2,1)\nprint *, m(1,3)\nend program t\n",
        ["1", "2", "5"]
    };
    reshape_fortran_3x2_last_row => {
        "program t\ninteger :: a(6) = [1, 2, 3, 4, 5, 6]\ninteger :: m(3,2)\nm = reshape(a, [3, 2])\nprint *, m(3,1)\nprint *, m(3,2)\nprint *, sum(m)\nend program t\n",
        ["5", "6", "21"]
    };
    reshape_fortran_2x2_diagonal => {
        "program t\ninteger :: a(4) = [1, 2, 3, 4]\ninteger :: m(2,2)\nm = reshape(a, [2, 2])\nprint *, m(1,1)\nprint *, m(2,2)\nprint *, m(2,1)\nend program t\n",
        ["1", "4", "2"]
    };
    reshape_fortran_1x6_row_vector => {
        "program t\ninteger :: a(6) = [10, 20, 30, 40, 50, 60]\ninteger :: m(1,6)\nm = reshape(a, [1, 6])\nprint *, m(1,1)\nprint *, m(1,6)\nend program t\n",
        ["10", "60"]
    };
    reshape_fortran_6x1_column_vector => {
        "program t\ninteger :: a(6) = [10, 20, 30, 40, 50, 60]\ninteger :: m(6,1)\nm = reshape(a, [6, 1])\nprint *, m(1,1)\nprint *, m(6,1)\nend program t\n",
        ["10", "60"]
    };
    reshape_fortran_3x3_center => {
        "program t\ninteger :: a(9) = [(i, i = 1, 9)]\ninteger :: m(3,3)\nm = reshape(a, [3, 3])\nprint *, m(2,2)\nprint *, m(3,3)\nend program t\n",
        ["5", "9"]
    };
    reshape_fortran_2x4_second_row => {
        "program t\ninteger :: a(8) = [1, 2, 3, 4, 5, 6, 7, 8]\ninteger :: m(2,4)\nm = reshape(a, [2, 4])\nprint *, m(1,2)\nprint *, m(2,2)\nprint *, m(2,4)\nend program t\n",
        ["3", "4", "8"]
    };
    reshape_fortran_4x2_corners => {
        "program t\ninteger :: a(8) = [1, 2, 3, 4, 5, 6, 7, 8]\ninteger :: m(4,2)\nm = reshape(a, [4, 2])\nprint *, m(1,1)\nprint *, m(4,1)\nprint *, m(4,2)\nend program t\n",
        ["1", "7", "8"]
    };
    reshape_fortran_3x2_sum_check => {
        "program t\ninteger :: a(6) = [2, 2, 2, 2, 2, 2]\ninteger :: m(3,2)\nm = reshape(a, [3, 2])\nprint *, sum(m)\nend program t\n",
        ["12"]
    };
    reshape_fortran_2x3_negative_values => {
        "program t\ninteger :: a(6) = [-1, 2, -3, 4, -5, 6]\ninteger :: m(2,3)\nm = reshape(a, [2, 3])\nprint *, m(1,1)\nprint *, m(2,3)\nend program t\n",
        ["-1", "6"]
    };

    // ── ORDER='C' row-major (10) ──────────────────────────────────────

    reshape_order_c_2x3_first_row => {
        "program t\ninteger :: a(6) = [1, 2, 3, 4, 5, 6]\ninteger :: m(2,3)\nm = reshape(a, [2, 3], order='C')\nprint *, m(1,1)\nprint *, m(1,3)\nprint *, m(2,1)\nend program t\n",
        ["1", "3", "4"]
    };
    reshape_order_c_3x2_last_column => {
        "program t\ninteger :: a(6) = [1, 2, 3, 4, 5, 6]\ninteger :: m(3,2)\nm = reshape(a, [3, 2], order='C')\nprint *, m(1,2)\nprint *, m(3,2)\nend program t\n",
        ["2", "6"]
    };
    reshape_order_c_2x2_anti_diagonal => {
        "program t\ninteger :: a(4) = [1, 2, 3, 4]\ninteger :: m(2,2)\nm = reshape(a, [2, 2], order='C')\nprint *, m(1,1)\nprint *, m(1,2)\nprint *, m(2,1)\nprint *, m(2,2)\nend program t\n",
        ["1", "2", "3", "4"]
    };
    reshape_order_c_1x4_linear => {
        "program t\ninteger :: a(4) = [9, 8, 7, 6]\ninteger :: m(1,4)\nm = reshape(a, [1, 4], order='C')\nprint *, m(1,1)\nprint *, m(1,4)\nend program t\n",
        ["9", "6"]
    };
    reshape_order_c_4x1_column => {
        "program t\ninteger :: a(4) = [9, 8, 7, 6]\ninteger :: m(4,1)\nm = reshape(a, [4, 1], order='C')\nprint *, m(1,1)\nprint *, m(4,1)\nend program t\n",
        ["9", "6"]
    };
    reshape_order_c_3x3_center => {
        "program t\ninteger :: a(9) = [(i, i = 1, 9)]\ninteger :: m(3,3)\nm = reshape(a, [3, 3], order='C')\nprint *, m(2,2)\nprint *, m(3,3)\nend program t\n",
        ["5", "9"]
    };
    reshape_order_c_2x4_corners => {
        "program t\ninteger :: a(8) = [1, 2, 3, 4, 5, 6, 7, 8]\ninteger :: m(2,4)\nm = reshape(a, [2, 4], order='C')\nprint *, m(1,1)\nprint *, m(1,4)\nprint *, m(2,4)\nend program t\n",
        ["1", "4", "8"]
    };
    reshape_order_c_4x2_first_column => {
        "program t\ninteger :: a(8) = [1, 2, 3, 4, 5, 6, 7, 8]\ninteger :: m(4,2)\nm = reshape(a, [4, 2], order='C')\nprint *, m(1,1)\nprint *, m(4,1)\nprint *, m(4,2)\nend program t\n",
        ["1", "7", "8"]
    };
    reshape_order_c_differs_from_fortran => {
        "program t\ninteger :: a(4) = [1, 2, 3, 4]\ninteger :: mf(2,2), mc(2,2)\nmf = reshape(a, [2, 2])\nmc = reshape(a, [2, 2], order='C')\nprint *, mf(1,2)\nprint *, mc(1,2)\nend program t\n",
        ["3", "2"]
    };
    reshape_order_c_sum_same_total => {
        "program t\ninteger :: a(6) = [1, 2, 3, 4, 5, 6]\ninteger :: mf(2,3), mc(2,3)\nmf = reshape(a, [2, 3])\nmc = reshape(a, [2, 3], order='C')\nprint *, sum(mf)\nprint *, sum(mc)\nend program t\n",
        ["21", "21"]
    };

    // ── PAD= constant fill for undersized source (12) ─────────────────

    reshape_pad_integer_zero_fill => {
        "program t\ninteger :: a(4) = [1, 2, 3, 4]\ninteger :: m(2,3)\nm = reshape(a, [2, 3], pad=0)\nprint *, m(1,3)\nprint *, m(2,3)\nprint *, sum(m)\nend program t\n",
        ["0", "0", "10"]
    };
    reshape_pad_integer_nine_fill => {
        "program t\ninteger :: a(3) = [1, 2, 3]\ninteger :: m(2,2)\nm = reshape(a, [2, 2], pad=9)\nprint *, m(1,2)\nprint *, m(2,1)\nprint *, m(2,2)\nend program t\n",
        ["2", "3", "9"]
    };
    reshape_pad_single_element_to_3 => {
        "program t\ninteger :: a(1) = [42]\ninteger :: m(3)\nm = reshape(a, [3], pad=-1)\nprint *, m(1)\nprint *, m(2)\nprint *, m(3)\nend program t\n",
        ["42", "-1", "-1"]
    };
    reshape_pad_2x2_from_three_elements => {
        "program t\ninteger :: a(3) = [10, 20, 30]\ninteger :: m(2,2)\nm = reshape(a, [2, 2], pad=0)\nprint *, m(1,1)\nprint *, m(2,2)\nend program t\n",
        ["10", "0"]
    };
    reshape_pad_3x3_from_five => {
        "program t\ninteger :: a(5) = [1, 2, 3, 4, 5]\ninteger :: m(3,3)\nm = reshape(a, [3, 3], pad=7)\nprint *, m(3,3)\nprint *, count(m == 7)\nend program t\n",
        ["7", "4"]
    };
    reshape_pad_with_order_c => {
        "program t\ninteger :: a(3) = [1, 2, 3]\ninteger :: m(2,2)\nm = reshape(a, [2, 2], pad=0, order='C')\nprint *, m(1,1)\nprint *, m(1,2)\nprint *, m(2,2)\nend program t\n",
        ["1", "2", "0"]
    };
    reshape_pad_real_half_fill => {
        "program t\nreal :: a(2) = [1.5, 2.5]\nreal :: m(2,2)\nm = reshape(a, [2, 2], pad=0.0)\nprint *, int(m(2,1) * 10)\nprint *, int(m(2,2) * 10)\nend program t\n",
        ["0", "0"]
    };
    reshape_pad_larger_than_source_2x4 => {
        "program t\ninteger :: a(5) = [1, 2, 3, 4, 5]\ninteger :: m(2,4)\nm = reshape(a, [2, 4], pad=0)\nprint *, m(2,4)\nprint *, sum(m)\nend program t\n",
        ["0", "15"]
    };
    reshape_pad_negative_pad_value => {
        "program t\ninteger :: a(2) = [5, 10]\ninteger :: m(2,2)\nm = reshape(a, [2, 2], pad=-99)\nprint *, m(1,2)\nprint *, m(2,1)\nend program t\n",
        ["-99", "10"]
    };
    reshape_pad_exact_fit_no_pad_used => {
        "program t\ninteger :: a(4) = [1, 2, 3, 4]\ninteger :: m(2,2)\nm = reshape(a, [2, 2], pad=99)\nprint *, sum(m)\nend program t\n",
        ["10"]
    };
    reshape_pad_1x5_from_two => {
        "program t\ninteger :: a(2) = [8, 9]\ninteger :: m(1,5)\nm = reshape(a, [1, 5], pad=0)\nprint *, m(1,1)\nprint *, m(1,5)\nend program t\n",
        ["8", "0"]
    };
    reshape_pad_5x1_from_two => {
        "program t\ninteger :: a(2) = [8, 9]\ninteger :: m(5,1)\nm = reshape(a, [5, 1], pad=0)\nprint *, m(1,1)\nprint *, m(5,1)\nend program t\n",
        ["8", "0"]
    };

    // ── 3D reshape totals (10) ────────────────────────────────────────

    reshape_3d_fortran_2x2x2_corner => {
        "program t\ninteger :: a(8) = [(i, i = 1, 8)]\ninteger :: m(2,2,2)\nm = reshape(a, [2, 2, 2])\nprint *, m(1,1,1)\nprint *, m(2,2,2)\nprint *, sum(m)\nend program t\n",
        ["1", "8", "36"]
    };
    reshape_3d_order_c_first_slice => {
        "program t\ninteger :: a(8) = [(i, i = 1, 8)]\ninteger :: m(2,2,2)\nm = reshape(a, [2, 2, 2], order='C')\nprint *, m(1,1,1)\nprint *, m(1,1,2)\nprint *, m(2,2,2)\nend program t\n",
        ["1", "2", "8"]
    };
    reshape_3d_pad_fill_last_layer => {
        "program t\ninteger :: a(5) = [1, 2, 3, 4, 5]\ninteger :: m(2,2,2)\nm = reshape(a, [2, 2, 2], pad=0)\nprint *, m(2,2,2)\nprint *, count(m == 0)\nend program t\n",
        ["0", "3"]
    };
    reshape_3d_2x3x1_column => {
        "program t\ninteger :: a(6) = [1, 2, 3, 4, 5, 6]\ninteger :: m(2,3,1)\nm = reshape(a, [2, 3, 1])\nprint *, m(2,3,1)\nprint *, sum(m)\nend program t\n",
        ["6", "21"]
    };
    reshape_3d_1x2x3_row => {
        "program t\ninteger :: a(6) = [1, 2, 3, 4, 5, 6]\ninteger :: m(1,2,3)\nm = reshape(a, [1, 2, 3])\nprint *, m(1,1,1)\nprint *, m(1,2,3)\nend program t\n",
        ["1", "6"]
    };
    reshape_3d_3x1x2 => {
        "program t\ninteger :: a(6) = [10, 20, 30, 40, 50, 60]\ninteger :: m(3,1,2)\nm = reshape(a, [3, 1, 2])\nprint *, m(1,1,1)\nprint *, m(3,1,2)\nend program t\n",
        ["10", "60"]
    };
    reshape_3d_order_c_differs => {
        "program t\ninteger :: a(4) = [1, 2, 3, 4]\ninteger :: mf(2,2,1), mc(2,2,1)\nmf = reshape(a, [2, 2, 1])\nmc = reshape(a, [2, 2, 1], order='C')\nprint *, mf(2,1,1)\nprint *, mc(2,1,1)\nend program t\n",
        ["2", "3"]
    };
    reshape_3d_pad_constant_layer => {
        "program t\ninteger :: a(3) = [1, 2, 3]\ninteger :: m(1,1,4)\nm = reshape(a, [1, 1, 4], pad=9)\nprint *, m(1,1,4)\nend program t\n",
        ["9"]
    };
    reshape_3d_sum_invariant => {
        "program t\ninteger :: a(12) = [(i, i = 1, 12)]\ninteger :: m(2,3,2)\nm = reshape(a, [2, 3, 2])\nprint *, sum(m)\nend program t\n",
        ["78"]
    };
    reshape_3d_order_c_sum_invariant => {
        "program t\ninteger :: a(12) = [(i, i = 1, 12)]\ninteger :: m(2,3,2)\nm = reshape(a, [2, 3, 2], order='C')\nprint *, sum(m)\nend program t\n",
        ["78"]
    };

    // ── Real and reshape expression forms (8) ─────────────────────────

    reshape_real_2x2_fractions => {
        "program t\nreal :: a(4) = [0.5, 1.5, 2.5, 3.5]\nreal :: m(2,2)\nm = reshape(a, [2, 2])\nprint *, int(sum(m) * 10)\nend program t\n",
        ["80"]
    };
    reshape_real_pad_zeros => {
        "program t\nreal :: a(2) = [1.0, 2.0]\nreal :: m(3)\nm = reshape(a, [3], pad=0.0)\nprint *, int(m(3) * 10)\nend program t\n",
        ["0"]
    };
    reshape_from_array_constructor => {
        "program t\ninteger :: m(2,2)\nm = reshape([(i, i = 1, 4)], [2, 2])\nprint *, m(2,2)\nend program t\n",
        ["4"]
    };
    reshape_into_existing_variable => {
        "program t\ninteger :: a(6) = [1, 2, 3, 4, 5, 6]\ninteger :: m(2,3)\nm = reshape(a, shape=[2, 3])\nprint *, m(1,2)\nend program t\n",
        ["3"]
    };
    reshape_shape_product_12 => {
        "program t\ninteger :: a(12) = [(i, i = 1, 12)]\ninteger :: m(3,4)\nm = reshape(a, [3, 4])\nprint *, m(3,4)\nprint *, sum(m)\nend program t\n",
        ["12", "78"]
    };
    reshape_shape_product_24_with_pad => {
        "program t\ninteger :: a(10) = [(i, i = 1, 10)]\ninteger :: m(2,3,4)\nm = reshape(a, [2, 3, 4], pad=0)\nprint *, sum(m)\nend program t\n",
        ["55"]
    };
    reshape_transpose_via_reshape_c => {
        "program t\ninteger :: a(2,3) = reshape([1, 4, 2, 5, 3, 6], [2, 3])\ninteger :: flat(6), back(3,2)\nflat = reshape(a, [6], order='C')\nback = reshape(flat, [3, 2], order='C')\nprint *, back(1,1)\nprint *, back(3,2)\nend program t\n",
        ["1", "6"]
    };
    reshape_pad_and_order_c_combined => {
        "program t\ninteger :: a(3) = [10, 20, 30]\ninteger :: m(2,3)\nm = reshape(a, [2, 3], pad=1, order='C')\nprint *, m(1,1)\nprint *, m(2,3)\nend program t\n",
        ["10", "1"]
    };

    // ── Edge totals and identity reshapes (10) ────────────────────────

    reshape_1d_to_1d_identity => {
        "program t\ninteger :: a(5) = [5, 4, 3, 2, 1]\ninteger :: b(5)\nb = reshape(a, [5])\nprint *, b(1)\nprint *, b(5)\nend program t\n",
        ["5", "1"]
    };
    reshape_2x2_from_transpose_source => {
        "program t\ninteger :: a(2,2) = reshape([1, 3, 2, 4], [2, 2])\ninteger :: b(2,2)\nb = reshape(a, [2, 2])\nprint *, b(1,1)\nprint *, b(2,2)\nend program t\n",
        ["1", "4"]
    };
    reshape_flatten_2d_to_1d => {
        "program t\ninteger :: a(2,3) = reshape([1, 2, 3, 4, 5, 6], [2, 3])\ninteger :: b(6)\nb = reshape(a, [6])\nprint *, b(1)\nprint *, b(6)\nend program t\n",
        ["1", "6"]
    };
    reshape_flatten_order_c => {
        "program t\ninteger :: a(2,2) = reshape([1, 2, 3, 4], [2, 2])\ninteger :: b(4)\nb = reshape(a, [4], order='C')\nprint *, b(1)\nprint *, b(4)\nend program t\n",
        ["1", "4"]
    };
    reshape_expand_1d_to_2d_count => {
        "program t\ninteger :: a(4) = [1, 1, 1, 1]\ninteger :: m(2,2)\nm = reshape(a, [2, 2])\nprint *, count(m == 1)\nend program t\n",
        ["4"]
    };
    reshape_all_zeros_pad_stays_zero => {
        "program t\ninteger :: a(2) = [0, 0]\ninteger :: m(2,2)\nm = reshape(a, [2, 2], pad=5)\nprint *, count(m == 0)\nprint *, count(m == 5)\nend program t\n",
        ["2", "2"]
    };
    reshape_large_pad_count_2x5 => {
        "program t\ninteger :: a(1) = [7]\ninteger :: m(2,5)\nm = reshape(a, [2, 5], pad=0)\nprint *, m(1,1)\nprint *, count(m == 0)\nend program t\n",
        ["7", "9"]
    };
    reshape_repeated_values_pattern => {
        "program t\ninteger :: a(8) = [2, 2, 2, 2, 2, 2, 2, 2]\ninteger :: m(2,2,2)\nm = reshape(a, [2, 2, 2])\nprint *, sum(m)\nend program t\n",
        ["16"]
    };
    reshape_mixed_sign_pad => {
        "program t\ninteger :: a(3) = [1, -1, 1]\ninteger :: m(2,2)\nm = reshape(a, [2, 2], pad=-1)\nprint *, m(2,2)\nend program t\n",
        ["-1"]
    };
    reshape_source_larger_than_shape_truncates => {
        "program t\ninteger :: a(10) = [(i, i = 1, 10)]\ninteger :: m(2,2)\nm = reshape(a, [2, 2])\nprint *, sum(m)\nend program t\n",
        ["10"]
    };
}

// ── Compile-only: invalid SHAPE (negative/zero size) ────────────────

#[test]
fn reshape_compile_negative_dimension() {
    compile_ok(
        r#"
program t
    integer :: a(4) = [1, 2, 3, 4]
    integer :: m(-1,2)
    m = reshape(a, [-1, 2])
    print *, m(1,1)
end program t
"#,
    );
}

#[test]
fn reshape_compile_zero_dimension() {
    compile_ok(
        r#"
program t
    integer :: a(4) = [1, 2, 3, 4]
    integer :: m(0,2)
    m = reshape(a, [0, 2])
    print *, 0
end program t
"#,
    );
}

#[test]
fn reshape_compile_shape_mismatch_no_pad() {
    compile_ok(
        r#"
program t
    integer :: a(3) = [1, 2, 3]
    integer :: m(2,2)
    m = reshape(a, [2, 2])
    print *, m(1,1)
end program t
"#,
    );
}

#[test]
fn reshape_compile_3d_negative_extent() {
    compile_ok(
        r#"
program t
    integer :: a(8) = [(i, i = 1, 8)]
    integer :: m(2,2,-1)
    m = reshape(a, [2, 2, -1])
    print *, 0
end program t
"#,
    );
}

#[test]
fn reshape_compile_order_invalid_literal() {
    compile_ok(
        r#"
program t
    integer :: a(4) = [1, 2, 3, 4]
    integer :: m(2,2)
    m = reshape(a, [2, 2], order='X')
    print *, m(1,1)
end program t
"#,
    );
}

#[test]
fn reshape_compile_pad_without_source() {
    compile_ok(
        r#"
program t
    integer :: m(3,3)
    m = reshape([(i, i = 1, 4)], [3, 3], pad=0)
    print *, m(3,3)
end program t
"#,
    );
}

#[test]
fn reshape_compile_empty_source_with_pad() {
    compile_ok(
        r#"
program t
    integer :: a(0)
    integer :: m(2)
    m = reshape(a, [2], pad=9)
    print *, m(1)
end program t
"#,
    );
}

#[test]
fn reshape_compile_large_shape_vector() {
    compile_ok(
        r#"
program t
    integer :: a(6) = [1, 2, 3, 4, 5, 6]
    integer :: sh(4) = [1, 1, 2, 3]
    integer :: m(1,1,2,3)
    m = reshape(a, sh)
    print *, sum(m)
end program t
"#,
    );
}
