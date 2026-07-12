//! Extended FORALL coverage: triple-index triplets, multi-index masks,
//! array-section LHS, and array-section RHS.
//! Distinct from `test_forall_advanced.rs` (scalar masks, 2D triangles, stride,
//! nested forall, statement form) and `test_arrays.rs` (basic 1D/2D forall).

fortran_cases! {
    // ── Triple-index FORALL ─────────────────────────────────────────

    forall_3d_fill_cube => {
        "program t\ninteger :: a(2,2,2)\na = 0\nforall (i = 1:2, j = 1:2, k = 1:2)\na(i,j,k) = i * 100 + j * 10 + k\nend forall\nprint *, a(1,1,1)\nprint *, a(2,1,2)\nprint *, a(2,2,2)\nend program t\n",
        ["111", "212", "222"]
    };
    forall_3d_mask_sum_indices_leq => {
        "program t\ninteger :: a(3,3,3)\na = 0\nforall (i = 1:3, j = 1:3, k = 1:3, i + j + k <= 4)\na(i,j,k) = i + j + k\nend forall\nprint *, a(1,1,1)\nprint *, a(2,1,1)\nprint *, a(2,2,2)\nend program t\n",
        ["3", "4", "0"]
    };
    forall_3d_mask_ordered_indices => {
        "program t\ninteger :: a(3,3,3)\na = 0\nforall (i = 1:3, j = 1:3, k = 1:3, i <= j .and. j <= k)\na(i,j,k) = 1\nend forall\nprint *, a(1,1,1)\nprint *, a(1,2,3)\nprint *, a(2,1,3)\nend program t\n",
        ["1", "1", "0"]
    };
    forall_3d_stride_subcube => {
        "program t\ninteger :: a(4,4,4)\na = 0\nforall (i = 1:4:2, j = 1:4:2, k = 1:4:2)\na(i,j,k) = i * j * k\nend forall\nprint *, a(1,1,1)\nprint *, a(3,3,3)\nprint *, a(2,2,2)\nend program t\n",
        ["1", "27", "0"]
    };

    // ── FORALL mask on multiple indices ───────────────────────────────

    forall_mask_i_plus_j_equals => {
        "program t\ninteger :: m(4,4)\nm = 0\nforall (i = 1:4, j = 1:4, i + j == 5)\nm(i,j) = i * j\nend forall\nprint *, m(1,4)\nprint *, m(2,3)\nprint *, m(1,1)\nend program t\n",
        ["4", "6", "0"]
    };
    forall_mask_i_times_j_gt => {
        "program t\ninteger :: m(4,4)\nm = 0\nforall (i = 1:4, j = 1:4, i * j > 6)\nm(i,j) = i + j\nend forall\nprint *, m(2,4)\nprint *, m(3,3)\nprint *, m(2,2)\nend program t\n",
        ["6", "6", "0"]
    };
    forall_3d_mask_product_mod_three => {
        "program t\ninteger :: a(3,3,3)\na = 0\nforall (i = 1:3, j = 1:3, k = 1:3, mod(i * j * k, 3) == 0)\na(i,j,k) = i + j + k\nend forall\nprint *, a(1,1,3)\nprint *, a(2,3,1)\nprint *, a(2,2,2)\nend program t\n",
        ["5", "6", "0"]
    };
    forall_mask_i_le_j_and_j_lt_k => {
        "program t\ninteger :: a(3,3,3)\na = 0\nforall (i = 1:3, j = 1:3, k = 1:3, i <= j .and. j < k)\na(i,j,k) = 10 * i + j + k\nend forall\nprint *, a(1,1,2)\nprint *, a(1,2,3)\nprint *, a(2,2,2)\nend program t\n",
        ["13", "16", "0"]
    };

    // ── FORALL assignment to array section (LHS) ────────────────────

    forall_lhs_row_section => {
        "program t\ninteger :: m(3,4)\nm = 0\nforall (i = 1:3)\nm(i, 1:4) = i * 10\nend forall\nprint *, m(1,1)\nprint *, m(2,1)\nprint *, m(3,4)\nend program t\n",
        ["10", "20", "30"]
    };
    forall_lhs_col_section => {
        "program t\ninteger :: m(4,3)\nm = 0\nforall (j = 1:3)\nm(1:4, j) = j\nend forall\nprint *, m(1,2)\nprint *, m(4,2)\nprint *, m(2,1)\nend program t\n",
        ["2", "2", "1"]
    };
    forall_lhs_paired_element_sections => {
        "program t\ninteger :: a(8)\na = 0\nforall (i = 1:4)\na(2 * i - 1:2 * i) = i\nend forall\nprint *, a(1)\nprint *, a(2)\nprint *, a(7)\nprint *, a(8)\nend program t\n",
        ["1", "1", "4", "4"]
    };
    forall_3d_lhs_plane_section => {
        "program t\ninteger :: m(2,3,2)\nm = 0\nforall (k = 1:2, i = 1:2)\nm(i, 1:3, k) = i * 10 + k\nend forall\nprint *, m(1,2,1)\nprint *, m(2,3,2)\nprint *, m(1,1,2)\nend program t\n",
        ["11", "23", "12"]
    };

    // ── FORALL with array RHS (array section) ───────────────────────

    forall_array_rhs_row_to_matrix => {
        "program t\ninteger :: src(4) = [5, 6, 7, 8]\ninteger :: m(3,4)\nm = 0\nforall (i = 1:3)\nm(i, 1:4) = src(1:4)\nend forall\nprint *, m(1,3)\nprint *, m(2,4)\nprint *, m(3,1)\nend program t\n",
        ["7", "8", "5"]
    };
    forall_array_rhs_scaled_row => {
        "program t\ninteger :: u(4) = [1, 2, 3, 4]\ninteger :: m(4,4)\nm = 0\nforall (i = 1:4)\nm(i, 1:4) = u(1:4) * i\nend forall\nprint *, m(2,3)\nprint *, m(4,1)\nprint *, m(1,4)\nend program t\n",
        ["6", "4", "4"]
    };
    forall_array_rhs_col_broadcast => {
        "program t\ninteger :: v(3) = [1, 2, 3]\ninteger :: m(3,3)\nm = 0\nforall (j = 1:3)\nm(1:3, j) = v(1:3)\nend forall\nprint *, m(2,2)\nprint *, m(3,1)\nprint *, m(1,3)\nend program t\n",
        ["2", "3", "1"]
    };
    forall_array_rhs_whole_vector_copy => {
        "program t\ninteger :: src(4) = [10, 20, 30, 40]\ninteger :: dst(4) = 0\nforall (i = 1:1)\ndst(1:4) = src(1:4)\nend forall\nprint *, dst(1)\nprint *, dst(3)\nprint *, sum(dst)\nend program t\n",
        ["10", "30", "100"]
    };
}
