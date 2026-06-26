//! Extended Fortran complex arithmetic: cmplx, real/aimag, conjg, + - * /, abs,
//! comparisons via real parts, and complex array element access.

fortran_cases! {
    cmplx_7_8_runtime_parts => {
        "program t\ncomplex :: z\nz = cmplx(7.0, 8.0)\nprint *, nint(real(z))\nprint *, nint(aimag(z))\nend program t\n",
        ["7", "8"]
    };

    cmplx_one_arg_6_runtime_real => {
        "program t\ncomplex :: z\nz = cmplx(6.0)\nprint *, nint(real(z))\nprint *, nint(aimag(z))\nend program t\n",
        ["6", "0"]
    };

    cmplx_from_integers_9_1 => {
        "program t\ninteger :: i = 9, j = 1\ncomplex :: z\nz = cmplx(i, j)\nprint *, nint(real(z))\nprint *, nint(aimag(z))\nend program t\n",
        ["9", "1"]
    };

    cmplx_pure_imag_0_5 => {
        "program t\ncomplex :: z\nz = cmplx(0.0, 5.0)\nprint *, nint(real(z))\nprint *, nint(aimag(z))\nend program t\n",
        ["0", "5"]
    };

    cmplx_negative_real_positive_imag => {
        "program t\ncomplex :: z\nz = cmplx(-5.0, 3.0)\nprint *, nint(real(z))\nprint *, nint(aimag(z))\nend program t\n",
        ["-5", "3"]
    };

    cmplx_expression_args => {
        "program t\ncomplex :: z\nz = cmplx(1.0 + 4.0, 2.0 + 2.0)\nprint *, nint(real(z))\nprint *, nint(aimag(z))\nend program t\n",
        ["5", "4"]
    };

    cmplx_kind8_parts_nint => {
        "program t\ninteger, parameter :: dp = kind(1.0d0)\ncomplex(dp) :: z\nz = cmplx(11.0_dp, 13.0_dp, dp)\nprint *, nint(real(z))\nprint *, nint(aimag(z))\nend program t\n",
        ["11", "13"]
    };

    real_literal_9_2 => {
        "program t\ncomplex :: z = (9.0, 2.0)\nprint *, nint(real(z))\nend program t\n",
        ["9"]
    };

    aimag_literal_neg1_7 => {
        "program t\ncomplex :: z = (-1.0, 7.0)\nprint *, nint(aimag(z))\nend program t\n",
        ["7"]
    };

    real_of_sum_runtime => {
        "program t\ncomplex :: a, b, c\na = cmplx(2.0, 3.0)\nb = cmplx(4.0, 5.0)\nc = a + b\nprint *, nint(real(c))\nend program t\n",
        ["6"]
    };

    aimag_of_negated_complex => {
        "program t\ncomplex :: z, n\nz = cmplx(4.0, -6.0)\nn = -z\nprint *, nint(aimag(n))\nend program t\n",
        ["6"]
    };

    real_of_conjg_part => {
        "program t\ncomplex :: z, c\nz = cmplx(8.0, -3.0)\nc = conjg(z)\nprint *, nint(real(c))\nend program t\n",
        ["8"]
    };

    aimag_of_product_nint => {
        "program t\ncomplex :: a, b, p\na = cmplx(2.0, 0.0)\nb = cmplx(0.0, 3.0)\np = a * b\nprint *, nint(aimag(p))\nend program t\n",
        ["6"]
    };

    conjg_5_neg3_parts => {
        "program t\ncomplex :: z, c\nz = cmplx(5.0, -3.0)\nc = conjg(z)\nprint *, nint(real(c))\nprint *, nint(aimag(c))\nend program t\n",
        ["5", "3"]
    };

    conjg_neg2_7_parts => {
        "program t\ncomplex :: z, c\nz = cmplx(-2.0, 7.0)\nc = conjg(z)\nprint *, nint(real(c))\nprint *, nint(aimag(c))\nend program t\n",
        ["-2", "-7"]
    };

    conjg_of_added_complex => {
        "program t\ncomplex :: a, b, s, c\na = cmplx(1.0, 2.0)\nb = cmplx(3.0, 4.0)\ns = a + b\nc = conjg(s)\nprint *, nint(real(c))\nprint *, nint(aimag(c))\nend program t\n",
        ["4", "-6"]
    };

    conjg_twice_restores_parts => {
        "program t\ncomplex :: z, c, r\nz = cmplx(7.0, -2.0)\nc = conjg(z)\nr = conjg(c)\nprint *, nint(real(r))\nprint *, nint(aimag(r))\nend program t\n",
        ["7", "-2"]
    };

    add_12_34_parts_nint => {
        "program t\ncomplex :: a, b, c\na = cmplx(1.0, 2.0)\nb = cmplx(3.0, 4.0)\nc = a + b\nprint *, nint(real(c))\nprint *, nint(aimag(c))\nend program t\n",
        ["4", "6"]
    };

    add_unit_imag_sum_nint => {
        "program t\ncomplex :: a, b, c\na = cmplx(0.0, 1.0)\nb = cmplx(0.0, 1.0)\nc = a + b\nprint *, nint(real(c))\nprint *, nint(aimag(c))\nend program t\n",
        ["0", "2"]
    };

    sub_105_32_parts_nint => {
        "program t\ncomplex :: a, b, c\na = cmplx(10.0, 5.0)\nb = cmplx(3.0, 2.0)\nc = a - b\nprint *, nint(real(c))\nprint *, nint(aimag(c))\nend program t\n",
        ["7", "3"]
    };

    sub_to_neg_parts => {
        "program t\ncomplex :: a, b, c\na = cmplx(0.0, 0.0)\nb = cmplx(1.0, 1.0)\nc = a - b\nprint *, nint(real(c))\nprint *, nint(aimag(c))\nend program t\n",
        ["-1", "-1"]
    };

    add_real_scalar_to_complex_nint => {
        "program t\ncomplex :: z, r\nz = cmplx(2.0, 3.0)\nr = 1.0 + z\nprint *, nint(real(r))\nprint *, nint(aimag(r))\nend program t\n",
        ["3", "3"]
    };

    sub_real_scalar_from_complex_nint => {
        "program t\ncomplex :: z, r\nz = cmplx(5.0, 7.0)\nr = z - 2.0\nprint *, nint(real(r))\nprint *, nint(aimag(r))\nend program t\n",
        ["3", "7"]
    };

    mul_23_14_parts_nint => {
        "program t\ncomplex :: a, b, c\na = cmplx(2.0, 3.0)\nb = cmplx(1.0, 4.0)\nc = a * b\nprint *, nint(real(c))\nprint *, nint(aimag(c))\nend program t\n",
        ["-10", "11"]
    };

    mul_real_by_imag_unit_nint => {
        "program t\ncomplex :: a, b, c\na = cmplx(1.0, 0.0)\nb = cmplx(0.0, 1.0)\nc = a * b\nprint *, nint(real(c))\nprint *, nint(aimag(c))\nend program t\n",
        ["0", "1"]
    };

    mul_pure_real_imag_nint => {
        "program t\ncomplex :: a, b, c\na = cmplx(3.0, 0.0)\nb = cmplx(0.0, 4.0)\nc = a * b\nprint *, nint(real(c))\nprint *, nint(aimag(c))\nend program t\n",
        ["0", "12"]
    };

    div_60_20_parts_nint => {
        "program t\ncomplex :: a, b, c\na = cmplx(6.0, 0.0)\nb = cmplx(2.0, 0.0)\nc = a / b\nprint *, nint(real(c))\nprint *, nint(aimag(c))\nend program t\n",
        ["3", "0"]
    };

    div_imag_over_imag_nint => {
        "program t\ncomplex :: a, b, c\na = cmplx(0.0, 8.0)\nb = cmplx(0.0, 2.0)\nc = a / b\nprint *, nint(real(c))\nprint *, nint(aimag(c))\nend program t\n",
        ["4", "0"]
    };

    div_34_by_i_parts_nint => {
        "program t\ncomplex :: a, b, c\na = cmplx(3.0, 4.0)\nb = cmplx(0.0, 1.0)\nc = a / b\nprint *, nint(real(c))\nprint *, nint(aimag(c))\nend program t\n",
        ["4", "-3"]
    };

    negate_23_parts_nint => {
        "program t\ncomplex :: z, n\nz = cmplx(2.0, 3.0)\nn = -z\nprint *, nint(real(n))\nprint *, nint(aimag(n))\nend program t\n",
        ["-2", "-3"]
    };

    abs_34_nint => {
        "program t\ncomplex :: z\nz = cmplx(3.0, 4.0)\nprint *, nint(abs(z))\nend program t\n",
        ["5"]
    };

    abs_512_nint => {
        "program t\ncomplex :: z\nz = cmplx(5.0, 12.0)\nprint *, nint(abs(z))\nend program t\n",
        ["13"]
    };

    abs_68_nint => {
        "program t\ncomplex :: z\nz = cmplx(6.0, 8.0)\nprint *, nint(abs(z))\nend program t\n",
        ["10"]
    };

    abs_neg34_nint => {
        "program t\ncomplex :: z\nz = cmplx(-3.0, -4.0)\nprint *, nint(abs(z))\nend program t\n",
        ["5"]
    };

    abs_z_plus_conjg_half_nint => {
        "program t\ncomplex :: z, s\nz = cmplx(3.0, 4.0)\ns = z + conjg(z)\nprint *, nint(abs(s) / 2.0)\nend program t\n",
        ["3"]
    };

    real_parts_equal_same_real_diff_imag => {
        "program t\ncomplex :: a = (2.0, 9.0), b = (2.0, 1.0)\nprint *, merge(1, 0, real(a) == real(b))\nend program t\n",
        ["1"]
    };

    real_parts_not_equal => {
        "program t\ncomplex :: a = (2.0, 0.0), b = (3.0, 0.0)\nprint *, merge(1, 0, real(a) /= real(b))\nend program t\n",
        ["1"]
    };

    real_part_less_than => {
        "program t\ncomplex :: a = (1.0, 5.0), b = (4.0, 1.0)\nprint *, merge(1, 0, real(a) < real(b))\nend program t\n",
        ["1"]
    };

    real_part_greater_than => {
        "program t\ncomplex :: a = (7.0, 2.0), b = (2.0, 7.0)\nprint *, merge(1, 0, real(a) > real(b))\nend program t\n",
        ["1"]
    };

    real_part_le_equal => {
        "program t\ncomplex :: a = (5.0, 1.0), b = (5.0, 9.0)\nprint *, merge(1, 0, real(a) <= real(b))\nend program t\n",
        ["1"]
    };

    real_part_ge_equal => {
        "program t\ncomplex :: a = (8.0, 3.0), b = (6.0, 8.0)\nprint *, merge(1, 0, real(a) >= real(b))\nend program t\n",
        ["1"]
    };

    array_element_real_index_1 => {
        "program t\ncomplex :: x(3)\nx(1) = cmplx(11.0, 0.0)\nx(2) = cmplx(22.0, 0.0)\nx(3) = cmplx(33.0, 0.0)\nprint *, nint(real(x(1)))\nend program t\n",
        ["11"]
    };

    array_element_aimag_index_4 => {
        "program t\ncomplex :: x(4)\nx(1) = cmplx(0.0, 1.0)\nx(2) = cmplx(0.0, 2.0)\nx(3) = cmplx(0.0, 3.0)\nx(4) = cmplx(0.0, 4.0)\nprint *, nint(aimag(x(4)))\nend program t\n",
        ["4"]
    };

    array_assign_cmplx_read_index_2 => {
        "program t\ncomplex :: x(3)\nx(2) = cmplx(11.0, 12.0)\nprint *, nint(real(x(2)))\nprint *, nint(aimag(x(2)))\nend program t\n",
        ["11", "12"]
    };

    array_elements_add_parts => {
        "program t\ncomplex :: x(2), y(2), z\nx(1) = cmplx(1.0, 2.0)\nx(2) = cmplx(3.0, 4.0)\ny(1) = cmplx(5.0, 6.0)\ny(2) = cmplx(7.0, 8.0)\nz = x(1) + y(2)\nprint *, nint(real(z))\nprint *, nint(aimag(z))\nend program t\n",
        ["8", "10"]
    };

    array_element_abs_nint => {
        "program t\ncomplex :: x(2)\nx(1) = cmplx(5.0, 12.0)\nx(2) = cmplx(1.0, 1.0)\nprint *, nint(abs(x(1)))\nend program t\n",
        ["13"]
    };

    array_loop_sum_real_parts => {
        "program t\ncomplex :: x(3)\ninteger :: i\nreal :: s\nx(1) = cmplx(1.0, 0.0)\nx(2) = cmplx(2.0, 0.0)\nx(3) = cmplx(3.0, 0.0)\ns = 0.0\ndo i = 1, 3\n  s = s + real(x(i))\nend do\nprint *, nint(s)\nend program t\n",
        ["6"]
    };

    array_compare_real_parts_merge => {
        "program t\ncomplex :: x(2)\nx(1) = cmplx(4.0, 1.0)\nx(2) = cmplx(4.0, 9.0)\nprint *, merge(1, 0, real(x(1)) == real(x(2)))\nend program t\n",
        ["1"]
    };

    array_2d_element_real_part => {
        "program t\ncomplex :: m(2, 2)\nm(1, 1) = cmplx(1.0, 2.0)\nm(2, 1) = cmplx(5.0, 6.0)\nprint *, nint(real(m(2, 1)))\nend program t\n",
        ["5"]
    };

    array_slice_second_real_part => {
        "program t\ncomplex :: a(4)\na(1) = cmplx(10.0, 0.0)\na(2) = cmplx(20.0, 0.0)\na(3) = cmplx(30.0, 0.0)\na(4) = cmplx(40.0, 0.0)\nprint *, nint(real(a(2:4)(2)))\nend program t\n",
        ["30"]
    };
}
