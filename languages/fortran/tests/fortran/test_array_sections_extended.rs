//! Extended array section coverage: bounded ranges, strides, whole sections,
//! variable bounds, multidimensional slices, section assignment, and reductions.
//! Distinct from `test_arrays.rs` (basic a(2:4), a(1:6:2), a(:3), a(3:), a(2:3,2:3)).

fortran_cases! {
    // ── Bounded range sections a(lo:hi) (8) ─────────────────────────

    section_2_to_5_sum_on_eight_elements => {
        "program t\ninteger :: a(8) = [1,2,3,4,5,6,7,8]\nprint *, sum(a(2:5))\nend program t\n",
        ["14"]
    };
    section_2_to_5_copy_first_and_last => {
        "program t\ninteger :: a(8) = [1,2,3,4,5,6,7,8]\ninteger :: b(4)\nb = a(2:5)\nprint *, b(1)\nprint *, b(4)\nend program t\n",
        ["2", "5"]
    };
    section_3_to_7_sum_ten_elements => {
        "program t\ninteger :: a(10) = [(i, i = 1, 10)]\nprint *, sum(a(3:7))\nend program t\n",
        ["25"]
    };
    section_1_to_6_corners_on_twelve => {
        "program t\ninteger :: a(12) = [(i, i = 1, 12)]\nprint *, a(1:6)(1)\nprint *, a(1:6)(6)\nprint *, size(a(1:6))\nend program t\n",
        ["1", "6", "6"]
    };
    section_5_to_8_four_element_window => {
        "program t\ninteger :: a(10) = [(i * 10, i = 1, 10)]\nprint *, a(5:8)(1)\nprint *, a(5:8)(4)\nprint *, sum(a(5:8))\nend program t\n",
        ["50", "80", "260"]
    };
    section_4_to_9_six_elements => {
        "program t\ninteger :: a(12) = [(i, i = 1, 12)]\nprint *, size(a(4:9))\nprint *, sum(a(4:9))\nend program t\n",
        ["6", "39"]
    };
    section_2_to_3_pair_sum => {
        "program t\ninteger :: a(5) = [10, 20, 30, 40, 50]\nprint *, sum(a(2:3))\nend program t\n",
        ["50"]
    };
    section_6_to_10_on_fifteen => {
        "program t\ninteger :: a(15) = [(i, i = 1, 15)]\nprint *, a(6:10)(1)\nprint *, a(6:10)(5)\nprint *, sum(a(6:10))\nend program t\n",
        ["6", "10", "40"]
    };

    // ── Stride sections a(lo:hi:step) (8) ───────────────────────────

    stride_1_to_10_by_2_sum => {
        "program t\ninteger :: a(10) = [(i, i = 1, 10)]\nprint *, sum(a(1:10:2))\nend program t\n",
        ["25"]
    };
    stride_1_to_9_by_3_sum => {
        "program t\ninteger :: a(9) = [(i, i = 1, 9)]\nprint *, sum(a(1:9:3))\nend program t\n",
        ["12"]
    };
    stride_2_to_10_by_2_sum => {
        "program t\ninteger :: a(10) = [(i, i = 1, 10)]\nprint *, sum(a(2:10:2))\nend program t\n",
        ["30"]
    };
    stride_1_to_10_by_2_copy_corners => {
        "program t\ninteger :: a(10) = [(i, i = 1, 10)]\ninteger :: b(5)\nb = a(1:10:2)\nprint *, b(1)\nprint *, b(5)\nprint *, size(b)\nend program t\n",
        ["1", "9", "5"]
    };
    stride_3_to_9_by_3_elements => {
        "program t\ninteger :: a(9) = [(i, i = 1, 9)]\nprint *, a(3:9:3)(1)\nprint *, a(3:9:3)(2)\nprint *, a(3:9:3)(3)\nend program t\n",
        ["3", "6", "9"]
    };
    stride_2_to_8_by_3_sum => {
        "program t\ninteger :: a(10) = [(i, i = 1, 10)]\nprint *, sum(a(2:8:3))\nend program t\n",
        ["15"]
    };
    stride_1_to_7_by_2_size => {
        "program t\ninteger :: a(7) = [(i, i = 1, 7)]\nprint *, size(a(1:7:2))\nprint *, sum(a(1:7:2))\nend program t\n",
        ["4", "16"]
    };
    stride_4_to_12_by_4_sum => {
        "program t\ninteger :: a(12) = [(i, i = 1, 12)]\nprint *, sum(a(4:12:4))\nend program t\n",
        ["24"]
    };

    // ── Whole array section a(:) (6) ──────────────────────────────────

    whole_colon_copy_matches_sum => {
        "program t\ninteger :: a(5) = [2, 4, 6, 8, 10]\ninteger :: b(5)\nb = a(:)\nprint *, sum(b)\nprint *, b(3)\nend program t\n",
        ["30", "6"]
    };
    whole_colon_print_first_last => {
        "program t\ninteger :: a(6) = [(i * 3, i = 1, 6)]\nprint *, a(:)(1)\nprint *, a(:)(6)\nend program t\n",
        ["3", "18"]
    };
    whole_colon_scalar_assign_zero => {
        "program t\ninteger :: a(4) = [9, 8, 7, 6]\na(:) = 0\nprint *, a(1)\nprint *, a(4)\nprint *, sum(a)\nend program t\n",
        ["0", "0", "0"]
    };
    whole_colon_double_in_place => {
        "program t\ninteger :: a(3) = [5, 10, 15]\na(:) = a(:) * 2\nprint *, a(1)\nprint *, a(3)\nprint *, sum(a)\nend program t\n",
        ["10", "30", "60"]
    };
    sum_whole_colon_section => {
        "program t\ninteger :: a(7) = [(i, i = 1, 7)]\nprint *, sum(a(:))\nend program t\n",
        ["28"]
    };
    whole_colon_assign_from_constructor => {
        "program t\ninteger :: a(4)\na(:) = [11, 22, 33, 44]\nprint *, a(2)\nprint *, sum(a(:))\nend program t\n",
        ["22", "110"]
    };

    // ── Variable-bound sections a(n:), a(:n), a(m:n) (6) ───────────

    section_n_to_end_with_n_four => {
        "program t\ninteger :: a(6) = [10, 20, 30, 40, 50, 60]\ninteger :: n\nn = 4\nprint *, sum(a(n:))\nprint *, size(a(n:))\nend program t\n",
        ["150", "3"]
    };
    section_1_to_n_with_n_five => {
        "program t\ninteger :: a(8) = [(i * 2, i = 1, 8)]\ninteger :: n\nn = 5\nprint *, a(1:n)(5)\nprint *, sum(a(1:n))\nend program t\n",
        ["10", "30"]
    };
    section_n_to_end_runtime_n_three => {
        "program t\ninteger :: a(5) = [1, 2, 3, 4, 5]\ninteger :: n\nn = 3\nprint *, a(n:)(1)\nprint *, a(n:)(3)\nend program t\n",
        ["3", "5"]
    };
    section_colon_n_first_three => {
        "program t\ninteger :: a(7) = [(i, i = 1, 7)]\ninteger :: n\nn = 3\nprint *, sum(a(:n))\nend program t\n",
        ["6"]
    };
    section_m_to_n_variable_bounds => {
        "program t\ninteger :: a(10) = [(i, i = 1, 10)]\ninteger :: m, n\nm = 2\nn = 6\nprint *, size(a(m:n))\nprint *, sum(a(m:n))\nend program t\n",
        ["5", "20"]
    };
    section_n_colon_with_stride => {
        "program t\ninteger :: a(10) = [(i, i = 1, 10)]\ninteger :: n\nn = 2\nprint *, sum(a(n:10:2))\nend program t\n",
        ["25"]
    };

    // ── Multidimensional sections (8) ───────────────────────────────

    section_2d_1_to_2_comma_2_to_3_sum => {
        "program t\ninteger :: a(3,4)\ninteger :: b(2,2)\ninteger :: i, j\ndo i = 1, 3\ndo j = 1, 4\na(i,j) = i * 10 + j\nend do\nend do\nb = a(1:2, 2:3)\nprint *, b(1,1)\nprint *, b(2,2)\nprint *, sum(b)\nend program t\n",
        ["12", "23", "70"]
    };
    section_2d_row_2_cols_1_to_4 => {
        "program t\ninteger :: a(3,4)\ninteger :: i, j\ndo i = 1, 3\ndo j = 1, 4\na(i,j) = i * 10 + j\nend do\nend do\nprint *, sum(a(2, 1:4))\nend program t\n",
        ["90"]
    };
    section_2d_col_3_rows_1_to_3 => {
        "program t\ninteger :: a(3,4)\ninteger :: i, j\ndo i = 1, 3\ndo j = 1, 4\na(i,j) = i * 10 + j\nend do\nend do\nprint *, a(1:3, 3)(1)\nprint *, a(1:3, 3)(3)\nprint *, sum(a(1:3, 3))\nend program t\n",
        ["13", "33", "69"]
    };
    section_2d_1_to_3_comma_1_to_2_size => {
        "program t\ninteger :: a(4,5)\ninteger :: b(3,2)\na = 1\nb = a(1:3, 1:2)\nprint *, size(b)\nprint *, sum(b)\nend program t\n",
        ["6", "6"]
    };
    section_2d_assign_submatrix => {
        "program t\ninteger :: a(3,3)\ninteger :: i, j\na = 0\ndo i = 1, 3\ndo j = 1, 3\na(i,j) = i + j\nend do\nend do\na(1:2, 2:3) = 0\nprint *, a(1,2)\nprint *, a(2,3)\nprint *, a(3,3)\nend program t\n",
        ["0", "0", "6"]
    };
    section_2d_row_slice_assign => {
        "program t\ninteger :: a(2,4)\na = 0\na(1, :) = [1, 2, 3, 4]\nprint *, a(1,1)\nprint *, a(1,4)\nprint *, sum(a(1,:))\nend program t\n",
        ["1", "4", "10"]
    };
    section_2d_col_slice_assign => {
        "program t\ninteger :: a(4,2)\na = 0\na(:, 2) = [5, 6, 7, 8]\nprint *, a(2,2)\nprint *, a(4,2)\nprint *, sum(a(:,2))\nend program t\n",
        ["6", "8", "26"]
    };
    section_2d_2_to_3_comma_colon_sum => {
        "program t\ninteger :: a(4,3)\ninteger :: i, j\ndo i = 1, 4\ndo j = 1, 3\na(i,j) = i + j\nend do\nend do\nprint *, sum(a(2:3, :))\nend program t\n",
        ["16"]
    };

    // ── Section assignment (7) ────────────────────────────────────────

    assign_section_2_to_5_scalar => {
        "program t\ninteger :: a(8) = [(i, i = 1, 8)]\na(2:5) = 7\nprint *, a(1)\nprint *, a(2)\nprint *, a(5)\nprint *, a(6)\nend program t\n",
        ["1", "7", "7", "6"]
    };
    assign_section_3_to_6_from_constructor => {
        "program t\ninteger :: a(8) = [(i, i = 1, 8)]\na(3:6) = [100, 200, 300, 400]\nprint *, a(2)\nprint *, a(3)\nprint *, a(6)\nprint *, a(7)\nend program t\n",
        ["2", "100", "400", "7"]
    };
    assign_section_stride_1_to_9_by_2 => {
        "program t\ninteger :: a(9) = [(i, i = 1, 9)]\na(1:9:2) = [11, 22, 33, 44, 55]\nprint *, a(1)\nprint *, a(3)\nprint *, a(9)\nprint *, sum(a)\nend program t\n",
        ["11", "22", "55", "189"]
    };
    assign_section_from_other_section => {
        "program t\ninteger :: a(6) = [6, 5, 4, 3, 2, 1]\ninteger :: b(3)\nb = a(2:4)\na(4:6) = b\nprint *, a(4)\nprint *, a(5)\nprint *, a(6)\nend program t\n",
        ["5", "4", "3"]
    };
    assign_section_2_to_4_increment => {
        "program t\ninteger :: a(5) = [1, 2, 3, 4, 5]\na(2:4) = a(2:4) + 10\nprint *, a(2)\nprint *, a(4)\nprint *, sum(a)\nend program t\n",
        ["12", "14", "45"]
    };
    assign_whole_colon_from_slice => {
        "program t\ninteger :: a(5) = [1, 2, 3, 4, 5]\ninteger :: b(3)\nb = [9, 8, 7]\na(:) = 0\na(2:4) = b\nprint *, a(1)\nprint *, a(3)\nprint *, a(5)\nend program t\n",
        ["0", "8", "0"]
    };
    assign_2d_section_from_constructor => {
        "program t\ninteger :: a(2,3)\na = 0\na(1:2, 2:3) = reshape([5, 6, 7, 8], [2, 2])\nprint *, a(1,2)\nprint *, a(2,3)\nprint *, sum(a)\nend program t\n",
        ["5", "8", "26"]
    };

    // ── Reductions and intrinsics on sections (7) ───────────────────

    sum_section_2_to_5_seven_elements => {
        "program t\ninteger :: a(7) = [3, 1, 9, 1, 5, 8, 2]\nprint *, sum(a(2:5))\nend program t\n",
        ["16"]
    };
    sum_section_stride_1_to_8_by_2 => {
        "program t\ninteger :: a(8) = [(i, i = 1, 8)]\nprint *, sum(a(1:8:2))\nend program t\n",
        ["16"]
    };
    sum_section_n_to_end => {
        "program t\ninteger :: a(6) = [2, 4, 6, 8, 10, 12]\ninteger :: n\nn = 3\nprint *, sum(a(n:))\nend program t\n",
        ["30"]
    };
    product_section_2_to_4 => {
        "program t\ninteger :: a(5) = [1, 2, 3, 4, 5]\nprint *, product(a(2:4))\nend program t\n",
        ["24"]
    };
    maxval_section_2_to_6 => {
        "program t\ninteger :: a(8) = [3, 1, 9, 1, 5, 8, 2, 7]\nprint *, maxval(a(2:6))\nend program t\n",
        ["9"]
    };
    minval_section_3_to_7 => {
        "program t\ninteger :: a(9) = [9, 8, 1, 4, 2, 7, 3, 6, 5]\nprint *, minval(a(3:7))\nend program t\n",
        ["1"]
    };
    size_section_2_to_5_and_stride => {
        "program t\ninteger :: a(10) = [(i, i = 1, 10)]\nprint *, size(a(2:5))\nprint *, size(a(1:10:2))\nend program t\n",
        ["4", "5"]
    };
}
