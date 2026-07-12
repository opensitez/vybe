//! Fortran array transforms: spread, cshift, eoshift, transpose, reshape, merge.
//! Distinct from `test_arrays_shift.rs` (compile-only shifts), `test_fortran2008.rs` (pack/unpack/spread compile),
//! and `test_arrays.rs` (basic reshape/transpose compile).

fortran_cases! {
    // ── SPREAD along dimension 2 (10) ─────────────────────────────
    spread_dim2_vector_123_copies_4_corner_and_sum => {
        "program t\ninteger :: a(3)=[1,2,3]\ninteger :: m(3,4)\nm=spread(a,2,4)\nprint *,m(1,1)\nprint *,m(2,4)\nprint *,sum(m)\nend program t\n",
        ["1", "2", "24"]
    };
    spread_dim2_singleton_five_copies_3 => {
        "program t\ninteger :: a(1)=[5]\ninteger :: m(1,3)\nm=spread(a,2,3)\nprint *,m(1,1)\nprint *,m(1,3)\nprint *,sum(m)\nend program t\n",
        ["5", "5", "15"]
    };
    spread_dim2_pair_values_copies_2 => {
        "program t\ninteger :: a(2)=[7,9]\ninteger :: m(2,2)\nm=spread(a,2,2)\nprint *,m(1,1)\nprint *,m(2,2)\nprint *,sum(m)\nend program t\n",
        ["7", "9", "32"]
    };
    spread_dim2_length4_copies_3_sum => {
        "program t\ninteger :: a(4)=[1,2,3,4]\ninteger :: m(4,3)\nm=spread(a,2,3)\nprint *,m(3,1)\nprint *,m(4,3)\nprint *,sum(m)\nend program t\n",
        ["3", "4", "30"]
    };
    spread_dim2_negatives_copies_2 => {
        "program t\ninteger :: a(3)=[-1,0,2]\ninteger :: m(3,2)\nm=spread(a,2,2)\nprint *,m(1,2)\nprint *,m(3,1)\nprint *,sum(m)\nend program t\n",
        ["-1", "2", "2"]
    };
    spread_dim2_ascending_copies_5 => {
        "program t\ninteger :: a(2)=[4,6]\ninteger :: m(2,5)\nm=spread(a,2,5)\nprint *,m(1,5)\nprint *,m(2,1)\nprint *,sum(m)\nend program t\n",
        ["4", "6", "50"]
    };
    spread_dim2_three_by_three_sum => {
        "program t\ninteger :: a(3)=[2,2,2]\ninteger :: m(3,3)\nm=spread(a,2,3)\nprint *,m(2,2)\nprint *,sum(m)\nend program t\n",
        ["2", "18"]
    };
    spread_dim2_four_copies_of_10_20 => {
        "program t\ninteger :: a(2)=[10,20]\ninteger :: m(2,4)\nm=spread(a,2,4)\nprint *,m(1,4)\nprint *,m(2,3)\nprint *,sum(m)\nend program t\n",
        ["10", "20", "120"]
    };
    spread_dim2_five_elements_two_copies => {
        "program t\ninteger :: a(5)=[1,1,2,2,3]\ninteger :: m(5,2)\nm=spread(a,2,2)\nprint *,m(5,1)\nprint *,m(1,2)\nprint *,sum(m)\nend program t\n",
        ["3", "1", "18"]
    };
    spread_dim2_zeros_copies_4 => {
        "program t\ninteger :: a(2)=[0,0]\ninteger :: m(2,4)\nm=spread(a,2,4)\nprint *,m(1,1)\nprint *,sum(m)\nend program t\n",
        ["0", "0"]
    };

    // ── SPREAD along dimension 1 (10) ─────────────────────────────
    spread_dim1_vector_123_copies_4 => {
        "program t\ninteger :: a(3)=[1,2,3]\ninteger :: m(4,3)\nm=spread(a,1,4)\nprint *,m(1,1)\nprint *,m(4,2)\nprint *,sum(m)\nend program t\n",
        ["1", "2", "24"]
    };
    spread_dim1_singleton_copies_3 => {
        "program t\ninteger :: a(1)=[8]\ninteger :: m(3,1)\nm=spread(a,1,3)\nprint *,m(1,1)\nprint *,m(3,1)\nprint *,sum(m)\nend program t\n",
        ["8", "8", "24"]
    };
    spread_dim1_pair_copies_2 => {
        "program t\ninteger :: a(2)=[3,5]\ninteger :: m(2,2)\nm=spread(a,1,2)\nprint *,m(1,1)\nprint *,m(2,2)\nprint *,sum(m)\nend program t\n",
        ["3", "5", "16"]
    };
    spread_dim1_length4_copies_3 => {
        "program t\ninteger :: a(4)=[1,2,3,4]\ninteger :: m(3,4)\nm=spread(a,1,3)\nprint *,m(2,3)\nprint *,m(3,4)\nprint *,sum(m)\nend program t\n",
        ["2", "4", "30"]
    };
    spread_dim1_negatives_copies_2 => {
        "program t\ninteger :: a(3)=[-2,1,4]\ninteger :: m(2,3)\nm=spread(a,1,2)\nprint *,m(1,2)\nprint *,m(2,3)\nprint *,sum(m)\nend program t\n",
        ["1", "4", "6"]
    };
    spread_dim1_ascending_copies_5 => {
        "program t\ninteger :: a(2)=[1,9]\ninteger :: m(5,2)\nm=spread(a,1,5)\nprint *,m(5,1)\nprint *,m(1,2)\nprint *,sum(m)\nend program t\n",
        ["1", "9", "50"]
    };
    spread_dim1_three_rows_sum => {
        "program t\ninteger :: a(3)=[4,5,6]\ninteger :: m(3,3)\nm=spread(a,1,3)\nprint *,m(3,2)\nprint *,sum(m)\nend program t\n",
        ["5", "45"]
    };
    spread_dim1_four_copies_wide => {
        "program t\ninteger :: a(2)=[11,13]\ninteger :: m(4,2)\nm=spread(a,1,4)\nprint *,m(4,1)\nprint *,m(2,2)\nprint *,sum(m)\nend program t\n",
        ["11", "13", "96"]
    };
    spread_dim1_five_by_two_sum => {
        "program t\ninteger :: a(2)=[2,3]\ninteger :: m(5,2)\nm=spread(a,1,5)\nprint *,m(5,2)\nprint *,sum(m)\nend program t\n",
        ["3", "25"]
    };
    spread_dim1_zeros_copies_3 => {
        "program t\ninteger :: a(1)=[0]\ninteger :: m(3,1)\nm=spread(a,1,3)\nprint *,m(2,1)\nprint *,sum(m)\nend program t\n",
        ["0", "0"]
    };

    // ── TRANSPOSE on small 2D matrices (8) ────────────────────────
    transpose_2x2_basic_corners_and_sum => {
        "program t\ninteger :: a(2,2)\ninteger :: b(2,2)\na(1,1)=1\na(1,2)=2\na(2,1)=3\na(2,2)=4\nb=transpose(a)\nprint *,b(1,1)\nprint *,b(2,1)\nprint *,sum(b)\nend program t\n",
        ["1", "2", "10"]
    };
    transpose_2x2_swap_off_diagonal => {
        "program t\ninteger :: a(2,2)\ninteger :: b(2,2)\na(1,1)=9\na(1,2)=8\na(2,1)=7\na(2,2)=6\nb=transpose(a)\nprint *,b(1,2)\nprint *,b(2,1)\nprint *,sum(b)\nend program t\n",
        ["7", "8", "30"]
    };
    transpose_2x3_row_matrix_sum => {
        "program t\ninteger :: a(3,2)\ninteger :: b(2,3)\na(1,1)=1\na(1,2)=2\na(2,1)=3\na(2,2)=4\na(3,1)=5\na(3,2)=6\nb=transpose(a)\nprint *,b(1,1)\nprint *,b(2,3)\nprint *,sum(b)\nend program t\n",
        ["1", "6", "21"]
    };
    transpose_3x2_column_layout => {
        "program t\ninteger :: a(2,3)\ninteger :: b(3,2)\na(1,1)=1\na(1,2)=2\na(1,3)=3\na(2,1)=4\na(2,2)=5\na(2,3)=6\nb=transpose(a)\nprint *,b(1,1)\nprint *,b(3,2)\nprint *,sum(b)\nend program t\n",
        ["1", "6", "21"]
    };
    transpose_2x3_all_nines => {
        "program t\ninteger :: a(3,2)\ninteger :: b(2,3)\na(1,1)=9\na(1,2)=9\na(2,1)=9\na(2,2)=9\na(3,1)=9\na(3,2)=9\nb=transpose(a)\nprint *,b(2,2)\nprint *,sum(b)\nend program t\n",
        ["9", "54"]
    };
    transpose_2x2_identity_like => {
        "program t\ninteger :: a(2,2)\ninteger :: b(2,2)\na(1,1)=1\na(1,2)=0\na(2,1)=0\na(2,2)=1\nb=transpose(a)\nprint *,b(1,1)\nprint *,b(2,2)\nprint *,sum(b)\nend program t\n",
        ["1", "1", "2"]
    };
    transpose_2x2_antidiagonal => {
        "program t\ninteger :: a(2,2)\ninteger :: b(2,2)\na(1,1)=0\na(1,2)=5\na(2,1)=7\na(2,2)=0\nb=transpose(a)\nprint *,b(1,2)\nprint *,b(2,1)\nprint *,sum(b)\nend program t\n",
        ["7", "5", "12"]
    };
    transpose_3x2_sequential_fill => {
        "program t\ninteger :: a(2,3)\ninteger :: b(3,2)\na(1,1)=10\na(1,2)=20\na(1,3)=30\na(2,1)=40\na(2,2)=50\na(2,3)=60\nb=transpose(a)\nprint *,b(1,1)\nprint *,b(3,2)\nprint *,sum(b)\nend program t\n",
        ["10", "60", "210"]
    };

    // ── MERGE scalar and array (8) ────────────────────────────────
    merge_scalar_true_picks_first => {
        "program t\ninteger :: x\nx=merge(42,99,.true.)\nprint *,x\nend program t\n",
        ["42"]
    };
    merge_scalar_false_picks_second => {
        "program t\ninteger :: x\nx=merge(42,99,.false.)\nprint *,x\nend program t\n",
        ["99"]
    };
    merge_scalar_negative_branch => {
        "program t\ninteger :: x\nx=merge(-3,7,.false.)\nprint *,x\nend program t\n",
        ["7"]
    };
    merge_scalar_zero_true_branch => {
        "program t\ninteger :: x\nx=merge(0,5,.true.)\nprint *,x\nend program t\n",
        ["0"]
    };
    merge_array_alternate_mask => {
        "program t\ninteger :: a(3)=[1,2,3]\ninteger :: b(3)=[4,5,6]\nlogical :: m(3)=[.true.,.false.,.true.]\ninteger :: c(3)\nc=merge(a,b,m)\nprint *,c(1)\nprint *,c(2)\nprint *,c(3)\nend program t\n",
        ["1", "5", "3"]
    };
    merge_array_all_true => {
        "program t\ninteger :: a(3)=[8,8,8]\ninteger :: b(3)=[1,2,3]\nlogical :: m(3)=[.true.,.true.,.true.]\ninteger :: c(3)\nc=merge(a,b,m)\nprint *,c(1)\nprint *,c(3)\nprint *,sum(c)\nend program t\n",
        ["8", "8", "24"]
    };
    merge_array_all_false => {
        "program t\ninteger :: a(3)=[8,8,8]\ninteger :: b(3)=[1,2,3]\nlogical :: m(3)=[.false.,.false.,.false.]\ninteger :: c(3)\nc=merge(a,b,m)\nprint *,c(1)\nprint *,c(2)\nprint *,c(3)\nend program t\n",
        ["1", "2", "3"]
    };
    merge_array_single_true_middle => {
        "program t\ninteger :: a(3)=[10,20,30]\ninteger :: b(3)=[1,2,3]\nlogical :: m(3)=[.false.,.true.,.false.]\ninteger :: c(3)\nc=merge(a,b,m)\nprint *,c(1)\nprint *,c(2)\nprint *,c(3)\nend program t\n",
        ["1", "20", "3"]
    };

    // ── CSHIFT 1D (6) ─────────────────────────────────────────────
    cshift_1d_left_one_first_and_fourth => {
        "program t\ninteger :: a(5)=[1,2,3,4,5]\ninteger :: b(5)\nb=cshift(a,1)\nprint *,b(1)\nprint *,b(4)\nend program t\n",
        ["2", "5"]
    };
    cshift_1d_left_two_first_second => {
        "program t\ninteger :: a(5)=[1,2,3,4,5]\ninteger :: b(5)\nb=cshift(a,2)\nprint *,b(1)\nprint *,b(2)\nend program t\n",
        ["3", "4"]
    };
    cshift_1d_right_one_first_fourth => {
        "program t\ninteger :: a(5)=[1,2,3,4,5]\ninteger :: b(5)\nb=cshift(a,-1)\nprint *,b(1)\nprint *,b(4)\nend program t\n",
        ["5", "3"]
    };
    cshift_1d_length4_rotate_three => {
        "program t\ninteger :: a(4)=[1,2,3,4]\ninteger :: b(4)\nb=cshift(a,3)\nprint *,b(1)\nprint *,b(4)\nend program t\n",
        ["2", "1"]
    };
    cshift_1d_zero_is_identity => {
        "program t\ninteger :: a(4)=[9,8,7,6]\ninteger :: b(4)\nb=cshift(a,0)\nprint *,b(1)\nprint *,b(4)\nend program t\n",
        ["9", "6"]
    };
    cshift_1d_negative_two_corners => {
        "program t\ninteger :: a(6)=[1,2,3,4,5,6]\ninteger :: b(6)\nb=cshift(a,-2)\nprint *,b(1)\nprint *,b(4)\nend program t\n",
        ["5", "2"]
    };

    // ── EOSHIFT 1D (6) ────────────────────────────────────────────
    eoshift_1d_left_one_first_last => {
        "program t\ninteger :: a(5)=[1,2,3,4,5]\ninteger :: b(5)\nb=eoshift(a,1)\nprint *,b(1)\nprint *,b(4)\nend program t\n",
        ["2", "5"]
    };
    eoshift_1d_right_one_fill_zero => {
        "program t\ninteger :: a(4)=[1,2,3,4]\ninteger :: b(4)\nb=eoshift(a,-1)\nprint *,b(1)\nprint *,b(4)\nend program t\n",
        ["0", "3"]
    };
    eoshift_1d_left_two_boundary => {
        "program t\ninteger :: a(5)=[1,2,3,4,5]\ninteger :: b(5)\nb=eoshift(a,2,-9)\nprint *,b(3)\nprint *,b(4)\nend program t\n",
        ["4", "5"]
    };
    eoshift_1d_zero_is_identity => {
        "program t\ninteger :: a(4)=[4,5,6,7]\ninteger :: b(4)\nb=eoshift(a,0)\nprint *,b(1)\nprint *,b(4)\nend program t\n",
        ["4", "7"]
    };
    eoshift_1d_left_one_length_three => {
        "program t\ninteger :: a(3)=[10,20,30]\ninteger :: b(3)\nb=eoshift(a,1)\nprint *,b(1)\nprint *,b(3)\nend program t\n",
        ["20", "30"]
    };
    eoshift_1d_right_two_boundary => {
        "program t\ninteger :: a(5)=[1,2,3,4,5]\ninteger :: b(5)\nb=eoshift(a,-2,0)\nprint *,b(1)\nprint *,b(3)\nend program t\n",
        ["0", "1"]
    };

    // ── RESHAPE default and order C (2) ───────────────────────────
    reshape_2x2_column_major_lower_row => {
        "program t\ninteger :: a(4)=[1,2,3,4]\ninteger :: m(2,2)\nm=reshape(a,[2,2])\nprint *,m(2,1)\nprint *,m(2,2)\nend program t\n",
        ["1", "2"]
    };
    reshape_2x2_order_c_upper_row => {
        "program t\ninteger :: a(4)=[1,2,3,4]\ninteger :: m(2,2)\nm=reshape(a,[2,2],'C')\nprint *,m(1,2)\nprint *,m(2,2)\nend program t\n",
        ["1", "2"]
    };
}
