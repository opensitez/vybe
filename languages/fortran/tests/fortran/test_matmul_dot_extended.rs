//! Extended MATMUL and DOT_PRODUCT coverage: distinct matrix shapes, identity
//! products, rectangular multiplies, vector-matrix combinations, slices, zeros,
//! and negative elements. Distinct from `test_arrays.rs` (basic 3-vector dot and
//! 2x2 times identity only).

fortran_cases! {
    // ── 2x2 general multiply (1) ────────────────────────────────────

    matmul_2x2_general_product => {
        "program t\ninteger :: a(2,2), b(2,2), c(2,2)\na(1,1)=1; a(1,2)=2; a(2,1)=3; a(2,2)=4\nb(1,1)=5; b(1,2)=6; b(2,1)=7; b(2,2)=8\nc = matmul(a, b)\nprint *, c(1,1)\nprint *, c(1,2)\nprint *, c(2,1)\nprint *, c(2,2)\nend program t\n",
        ["19", "22", "43", "50"]
    };

    // ── Rectangular shapes (3) ──────────────────────────────────────

    matmul_2x3_by_3x2_rectangular => {
        "program t\ninteger :: a(2,3), b(3,2), c(2,2)\na(1,1)=1; a(1,2)=2; a(1,3)=3; a(2,1)=4; a(2,2)=5; a(2,3)=6\nb(1,1)=7; b(1,2)=1; b(2,1)=8; b(2,2)=0; b(3,1)=9; b(3,2)=-1\nc = matmul(a, b)\nprint *, c(1,1)\nprint *, c(1,2)\nprint *, c(2,1)\nprint *, c(2,2)\nend program t\n",
        ["50", "-2", "122", "-2"]
    };
    matmul_3x2_by_2x3_expands_to_3x3 => {
        "program t\ninteger :: a(3,2), b(2,3), c(3,3)\na(1,1)=1; a(1,2)=2; a(2,1)=3; a(2,2)=4; a(3,1)=5; a(3,2)=6\nb(1,1)=1; b(1,2)=0; b(1,3)=-1; b(2,1)=0; b(2,2)=1; b(2,3)=0\nc = matmul(a, b)\nprint *, c(1,1)\nprint *, c(1,3)\nprint *, c(3,1)\nprint *, c(3,3)\nend program t\n",
        ["1", "-1", "5", "-5"]
    };
    matmul_1x3_by_3x1_scalar_shape => {
        "program t\ninteger :: a(1,3), b(3,1), c(1,1)\na(1,1)=1; a(1,2)=2; a(1,3)=3\nb(1,1)=4; b(2,1)=5; b(3,1)=6\nc = matmul(a, b)\nprint *, c(1,1)\nend program t\n",
        ["32"]
    };

    // ── Identity multiply (2) ───────────────────────────────────────

    matmul_identity_left_2x2 => {
        "program t\ninteger :: ident(2,2), a(2,2), c(2,2)\nident = 0; ident(1,1)=1; ident(2,2)=1\na(1,1)=7; a(1,2)=-3; a(2,1)=5; a(2,2)=2\nc = matmul(ident, a)\nprint *, c(1,1)\nprint *, c(1,2)\nprint *, c(2,1)\nprint *, c(2,2)\nend program t\n",
        ["7", "-3", "5", "2"]
    };
    matmul_identity_3x3_preserves => {
        "program t\ninteger :: ident(3,3), a(3,3), c(3,3)\nident = 0; ident(1,1)=1; ident(2,2)=1; ident(3,3)=1\na(1,1)=2; a(1,2)=3; a(1,3)=4; a(2,1)=5; a(2,2)=6; a(2,3)=7; a(3,1)=8; a(3,2)=9; a(3,3)=10\nc = matmul(a, ident)\nprint *, c(2,2)\nprint *, c(3,1)\nprint *, sum(c)\nend program t\n",
        ["6", "8", "54"]
    };

    // ── Vector-matrix combinations (2) ──────────────────────────────

    matmul_matrix_times_column_vector => {
        "program t\ninteger :: a(2,3), v(3), c(2)\na(1,1)=1; a(1,2)=2; a(1,3)=3; a(2,1)=4; a(2,2)=5; a(2,3)=6\nv = [1, 0, -1]\nc = matmul(a, v)\nprint *, c(1)\nprint *, c(2)\nend program t\n",
        ["-2", "-2"]
    };
    matmul_row_vector_times_matrix => {
        "program t\ninteger :: v(3), b(3,2), c(2)\nv = [1, 2, 3]\nb(1,1)=1; b(1,2)=0; b(2,1)=0; b(2,2)=1; b(3,1)=1; b(3,2)=1\nc = matmul(v, b)\nprint *, c(1)\nprint *, c(2)\nend program t\n",
        ["4", "5"]
    };

    // ── Zeros and negative elements in matmul (2) ───────────────────

    matmul_2x2_negative_mixed_signs => {
        "program t\ninteger :: a(2,2), b(2,2), c(2,2)\na(1,1)=1; a(1,2)=-1; a(2,1)=2; a(2,2)=-2\nb(1,1)=3; b(1,2)=4; b(2,1)=-1; b(2,2)=0\nc = matmul(a, b)\nprint *, c(1,1)\nprint *, c(1,2)\nprint *, c(2,1)\nprint *, c(2,2)\nend program t\n",
        ["4", "4", "8", "8"]
    };
    matmul_zero_matrix_2x2 => {
        "program t\ninteger :: z(2,2), b(2,2), c(2,2)\nz = 0\nb(1,1)=9; b(1,2)=-4; b(2,1)=3; b(2,2)=7\nc = matmul(z, b)\nprint *, c(1,1)\nprint *, c(1,2)\nprint *, c(2,1)\nprint *, c(2,2)\nend program t\n",
        ["0", "0", "0", "0"]
    };

    // ── DOT_PRODUCT on slices (2) ───────────────────────────────────

    dot_product_slice_bounded_range => {
        "program t\ninteger :: a(6) = [10, 20, 30, 40, 50, 60]\ninteger :: b(6) = [1, 1, 1, 1, 1, 1]\nprint *, dot_product(a(2:4), b(2:4))\nend program t\n",
        ["90"]
    };
    dot_product_slice_stride_two => {
        "program t\ninteger :: a(7) = [1, 2, 3, 4, 5, 6, 7]\ninteger :: b(7) = [7, 6, 5, 4, 3, 2, 1]\nprint *, dot_product(a(1:7:2), b(1:7:2))\nend program t\n",
        ["44"]
    };

    // ── Zero vectors (2) ────────────────────────────────────────────

    dot_product_zero_first_operand => {
        "program t\ninteger :: a(3) = [0, 0, 0]\ninteger :: b(3) = [2, 3, 4]\nprint *, dot_product(a, b)\nend program t\n",
        ["0"]
    };
    dot_product_both_zero_length_four => {
        "program t\ninteger :: a(4) = [0, 0, 0, 0]\ninteger :: b(4) = [0, 0, 0, 0]\nprint *, dot_product(a, b)\nend program t\n",
        ["0"]
    };

    // ── Negative and real dot products (2) ────────────────────────────

    dot_product_negative_mixed_signs => {
        "program t\ninteger :: a(3) = [-1, 2, -3]\ninteger :: b(3) = [4, -5, 6]\nprint *, dot_product(a, b)\nend program t\n",
        ["-32"]
    };
    dot_product_real_fractional_values => {
        "program t\nreal :: a(3) = [0.5, 1.5, 2.0]\nreal :: b(3) = [2.0, 2.0, 1.0]\nprint *, dot_product(a, b)\nend program t\n",
        ["6"]
    };
}
