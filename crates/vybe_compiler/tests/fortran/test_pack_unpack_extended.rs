//! Extended PACK/UNPACK: masked selection, VECTOR padding, 2D masks, real and
//! character arrays, round-trip restore. Distinct from `test_fortran2008.rs`
//! compile-only pack/unpack smokes and `test_array_transforms.rs` reshape/spread.

fortran_cases! {
    // ── Integer PACK without VECTOR (12) ────────────────────────────

    pack_int_alternating_mask_first_third_fifth => {
        "program t\ninteger :: a(5) = [10, 20, 30, 40, 50]\nlogical :: mask(5) = [.true., .false., .true., .false., .true.]\ninteger :: b(3)\nb = pack(a, mask)\nprint *, b(1)\nprint *, b(2)\nprint *, b(3)\nend program t\n",
        ["10", "30", "50"]
    };
    pack_int_leading_true_run => {
        "program t\ninteger :: a(6) = [1, 2, 3, 4, 5, 6]\nlogical :: mask(6) = [.true., .true., .true., .false., .false., .false.]\ninteger :: b(3)\nb = pack(a, mask)\nprint *, sum(b)\nend program t\n",
        ["6"]
    };
    pack_int_trailing_true_run => {
        "program t\ninteger :: a(6) = [1, 2, 3, 4, 5, 6]\nlogical :: mask(6) = [.false., .false., .false., .true., .true., .true.]\ninteger :: b(3)\nb = pack(a, mask)\nprint *, b(1)\nprint *, b(3)\nend program t\n",
        ["4", "6"]
    };
    pack_int_singleton_selected => {
        "program t\ninteger :: a(5) = [7, 8, 9, 10, 11]\nlogical :: mask(5) = [.false., .false., .true., .false., .false.]\ninteger :: b(1)\nb = pack(a, mask)\nprint *, b(1)\nend program t\n",
        ["9"]
    };
    pack_int_all_true_preserves_order => {
        "program t\ninteger :: a(4) = [3, 1, 4, 1]\nlogical :: mask(4) = [.true., .true., .true., .true.]\ninteger :: b(4)\nb = pack(a, mask)\nprint *, b(2)\nprint *, b(4)\nend program t\n",
        ["1", "1"]
    };
    pack_int_sparse_two_of_eight => {
        "program t\ninteger :: a(8) = [0, 0, 5, 0, 0, 0, 9, 0]\nlogical :: mask(8) = [.false., .false., .true., .false., .false., .false., .true., .false.]\ninteger :: b(2)\nb = pack(a, mask)\nprint *, b(1)\nprint *, b(2)\nend program t\n",
        ["5", "9"]
    };
    pack_int_negative_values_selected => {
        "program t\ninteger :: a(5) = [-3, 2, -7, 4, -1]\nlogical :: mask(5) = [.true., .false., .true., .false., .true.]\ninteger :: b(3)\nb = pack(a, mask)\nprint *, b(1)\nprint *, b(2)\nprint *, b(3)\nend program t\n",
        ["-3", "-7", "-1"]
    };
    pack_int_even_positions_only => {
        "program t\ninteger :: a(6) = [11, 22, 33, 44, 55, 66]\nlogical :: mask(6) = [.false., .true., .false., .true., .false., .true.]\ninteger :: b(3)\nb = pack(a, mask)\nprint *, sum(b)\nend program t\n",
        ["132"]
    };
    pack_int_odd_positions_only => {
        "program t\ninteger :: a(6) = [11, 22, 33, 44, 55, 66]\nlogical :: mask(6) = [.true., .false., .true., .false., .true., .false.]\ninteger :: b(3)\nb = pack(a, mask)\nprint *, sum(b)\nend program t\n",
        ["99"]
    };
    pack_int_interior_window => {
        "program t\ninteger :: a(7) = [100, 2, 3, 4, 5, 6, 200]\nlogical :: mask(7) = [.false., .true., .true., .true., .true., .true., .false.]\ninteger :: b(5)\nb = pack(a, mask)\nprint *, b(1)\nprint *, b(5)\nend program t\n",
        ["2", "6"]
    };
    pack_int_descending_selection => {
        "program t\ninteger :: a(5) = [9, 7, 5, 3, 1]\nlogical :: mask(5) = [.true., .true., .false., .true., .false.]\ninteger :: b(3)\nb = pack(a, mask)\nprint *, b(1)\nprint *, b(2)\nprint *, b(3)\nend program t\n",
        ["9", "7", "3"]
    };
    pack_int_zeros_in_source => {
        "program t\ninteger :: a(5) = [0, 0, 0, 0, 0]\nlogical :: mask(5) = [.true., .false., .true., .false., .true.]\ninteger :: b(3)\nb = pack(a, mask)\nprint *, count(b == 0)\nend program t\n",
        ["3"]
    };

    // ── Integer PACK with VECTOR padding (12) ───────────────────────

    pack_int_vector_pads_with_nines => {
        "program t\ninteger :: a(5) = [1, 2, 3, 4, 5]\nlogical :: mask(5) = [.true., .false., .true., .false., .true.]\ninteger :: vec(5) = [9, 9, 9, 9, 9]\ninteger :: b(5)\nb = pack(a, mask, vec)\nprint *, b(1)\nprint *, b(3)\nprint *, b(4)\nprint *, b(5)\nend program t\n",
        ["1", "5", "9", "9"]
    };
    pack_int_vector_length_four_from_three_selected => {
        "program t\ninteger :: a(6) = [2, 4, 6, 8, 10, 12]\nlogical :: mask(6) = [.true., .true., .true., .false., .false., .false.]\ninteger :: vec(4) = [-1, -1, -1, -1]\ninteger :: b(4)\nb = pack(a, mask, vec)\nprint *, b(1)\nprint *, b(3)\nprint *, b(4)\nend program t\n",
        ["2", "6", "-1"]
    };
    pack_int_vector_shorter_than_selection => {
        "program t\ninteger :: a(4) = [5, 6, 7, 8]\nlogical :: mask(4) = [.true., .true., .true., .true.]\ninteger :: vec(2) = [0, 0]\ninteger :: b(2)\nb = pack(a, mask, vec)\nprint *, b(1)\nprint *, b(2)\nend program t\n",
        ["5", "6"]
    };
    pack_int_vector_all_false_uses_vector => {
        "program t\ninteger :: a(3) = [1, 2, 3]\nlogical :: mask(3) = [.false., .false., .false.]\ninteger :: vec(3) = [88, 77, 66]\ninteger :: b(3)\nb = pack(a, mask, vec)\nprint *, b(1)\nprint *, b(2)\nprint *, b(3)\nend program t\n",
        ["88", "77", "66"]
    };
    pack_int_vector_single_pad_slot => {
        "program t\ninteger :: a(3) = [10, 20, 30]\nlogical :: mask(3) = [.true., .false., .false.]\ninteger :: vec(2) = [99, 99]\ninteger :: b(2)\nb = pack(a, mask, vec)\nprint *, b(1)\nprint *, b(2)\nend program t\n",
        ["10", "99"]
    };
    pack_int_vector_preserves_trailing_vector => {
        "program t\ninteger :: a(5) = [3, 6, 9, 12, 15]\nlogical :: mask(5) = [.false., .true., .false., .true., .false.]\ninteger :: vec(4) = [100, 200, 300, 400]\ninteger :: b(4)\nb = pack(a, mask, vec)\nprint *, b(1)\nprint *, b(2)\nprint *, b(3)\nend program t\n",
        ["6", "12", "300"]
    };
    pack_int_vector_two_selected_four_slots => {
        "program t\ninteger :: a(4) = [1, 3, 5, 7]\nlogical :: mask(4) = [.true., .false., .true., .false.]\ninteger :: vec(4) = [0, 0, 0, 0]\ninteger :: b(4)\nb = pack(a, mask, vec)\nprint *, b(1)\nprint *, b(2)\nprint *, b(4)\nend program t\n",
        ["1", "5", "0"]
    };
    pack_int_vector_negative_pad_values => {
        "program t\ninteger :: a(4) = [4, 8, 12, 16]\nlogical :: mask(4) = [.true., .true., .false., .false.]\ninteger :: vec(3) = [-5, -5, -5]\ninteger :: b(3)\nb = pack(a, mask, vec)\nprint *, b(1)\nprint *, b(2)\nprint *, b(3)\nend program t\n",
        ["4", "8", "-5"]
    };
    pack_int_vector_exact_fit_no_pad => {
        "program t\ninteger :: a(5) = [2, 4, 6, 8, 10]\nlogical :: mask(5) = [.true., .false., .true., .false., .true.]\ninteger :: vec(3) = [0, 0, 0]\ninteger :: b(3)\nb = pack(a, mask, vec)\nprint *, sum(b)\nend program t\n",
        ["18"]
    };
    pack_int_vector_longer_than_selection => {
        "program t\ninteger :: a(3) = [7, 14, 21]\nlogical :: mask(3) = [.true., .false., .true.]\ninteger :: vec(5) = [1, 2, 3, 4, 5]\ninteger :: b(5)\nb = pack(a, mask, vec)\nprint *, b(1)\nprint *, b(2)\nprint *, b(5)\nend program t\n",
        ["7", "21", "5"]
    };
    pack_int_vector_alternating_mask_six_slots => {
        "program t\ninteger :: a(6) = [1, 2, 3, 4, 5, 6]\nlogical :: mask(6) = [.true., .false., .true., .false., .true., .false.]\ninteger :: vec(6) = [99, 99, 99, 99, 99, 99]\ninteger :: b(6)\nb = pack(a, mask, vec)\nprint *, b(1)\nprint *, b(3)\nprint *, b(5)\nprint *, b(6)\nend program t\n",
        ["1", "3", "5", "99"]
    };
    pack_int_vector_first_last_only => {
        "program t\ninteger :: a(5) = [11, 22, 33, 44, 55]\nlogical :: mask(5) = [.true., .false., .false., .false., .true.]\ninteger :: vec(4) = [0, 0, 0, 0]\ninteger :: b(4)\nb = pack(a, mask, vec)\nprint *, b(1)\nprint *, b(2)\nprint *, b(3)\nend program t\n",
        ["11", "55", "0"]
    };

    // ── Integer UNPACK with FIELD fill (12) ───────────────────────────

    unpack_int_alternating_restore => {
        "program t\ninteger :: a(3) = [10, 30, 50]\nlogical :: mask(5) = [.true., .false., .true., .false., .true.]\ninteger :: fill(5) = [0, 0, 0, 0, 0]\ninteger :: b(5)\nb = unpack(a, mask, fill)\nprint *, b(1)\nprint *, b(2)\nprint *, b(5)\nend program t\n",
        ["10", "0", "50"]
    };
    unpack_int_custom_fill_value => {
        "program t\ninteger :: a(2) = [7, 9]\nlogical :: mask(4) = [.true., .false., .true., .false.]\ninteger :: fill(4) = [-1, -1, -1, -1]\ninteger :: b(4)\nb = unpack(a, mask, fill)\nprint *, b(1)\nprint *, b(2)\nprint *, b(3)\nend program t\n",
        ["7", "-1", "9"]
    };
    unpack_int_all_true_no_fill_used => {
        "program t\ninteger :: a(3) = [1, 2, 3]\nlogical :: mask(3) = [.true., .true., .true.]\ninteger :: fill(3) = [99, 99, 99]\ninteger :: b(3)\nb = unpack(a, mask, fill)\nprint *, sum(b)\nend program t\n",
        ["6"]
    };
    unpack_int_all_false_uses_fill => {
        "program t\ninteger :: a(1) = [42]\nlogical :: mask(3) = [.false., .false., .false.]\ninteger :: fill(3) = [5, 6, 7]\ninteger :: b(3)\nb = unpack(a, mask, fill)\nprint *, b(1)\nprint *, b(2)\nprint *, b(3)\nend program t\n",
        ["5", "6", "7"]
    };
    unpack_int_leading_run => {
        "program t\ninteger :: a(2) = [100, 200]\nlogical :: mask(4) = [.true., .true., .false., .false.]\ninteger :: fill(4) = [0, 0, 0, 0]\ninteger :: b(4)\nb = unpack(a, mask, fill)\nprint *, b(1)\nprint *, b(2)\nprint *, b(4)\nend program t\n",
        ["100", "200", "0"]
    };
    unpack_int_trailing_run => {
        "program t\ninteger :: a(2) = [3, 4]\nlogical :: mask(4) = [.false., .false., .true., .true.]\ninteger :: fill(4) = [1, 1, 1, 1]\ninteger :: b(4)\nb = unpack(a, mask, fill)\nprint *, b(1)\nprint *, b(3)\nprint *, b(4)\nend program t\n",
        ["1", "3", "4"]
    };
    unpack_int_singleton_into_middle => {
        "program t\ninteger :: a(1) = [42]\nlogical :: mask(5) = [.false., .false., .true., .false., .false.]\ninteger :: fill(5) = [0, 0, 0, 0, 0]\ninteger :: b(5)\nb = unpack(a, mask, fill)\nprint *, b(3)\nend program t\n",
        ["42"]
    };
    unpack_int_scattered_positions => {
        "program t\ninteger :: a(3) = [2, 4, 6]\nlogical :: mask(5) = [.true., .false., .true., .false., .true.]\ninteger :: fill(5) = [0, 0, 0, 0, 0]\ninteger :: b(5)\nb = unpack(a, mask, fill)\nprint *, b(1)\nprint *, b(3)\nprint *, b(5)\nend program t\n",
        ["2", "4", "6"]
    };
    unpack_int_negative_fill => {
        "program t\ninteger :: a(2) = [8, 16]\nlogical :: mask(4) = [.false., .true., .false., .true.]\ninteger :: fill(4) = [-9, -9, -9, -9]\ninteger :: b(4)\nb = unpack(a, mask, fill)\nprint *, b(1)\nprint *, b(2)\nprint *, b(4)\nend program t\n",
        ["-9", "8", "16"]
    };
    unpack_int_even_positions => {
        "program t\ninteger :: a(3) = [10, 20, 30]\nlogical :: mask(6) = [.false., .true., .false., .true., .false., .true.]\ninteger :: fill(6) = [0, 0, 0, 0, 0, 0]\ninteger :: b(6)\nb = unpack(a, mask, fill)\nprint *, b(2)\nprint *, b(4)\nprint *, b(6)\nend program t\n",
        ["10", "20", "30"]
    };
    unpack_int_odd_positions => {
        "program t\ninteger :: a(3) = [5, 15, 25]\nlogical :: mask(6) = [.true., .false., .true., .false., .true., .false.]\ninteger :: fill(6) = [0, 0, 0, 0, 0, 0]\ninteger :: b(6)\nb = unpack(a, mask, fill)\nprint *, b(1)\nprint *, b(3)\nprint *, b(5)\nend program t\n",
        ["5", "15", "25"]
    };
    unpack_int_fill_from_variable => {
        "program t\ninteger :: a(1) = [99]\nlogical :: mask(3) = [.false., .true., .false.]\ninteger :: fill(3) = [1, 2, 3]\ninteger :: b(3)\nb = unpack(a, mask, fill)\nprint *, b(2)\nend program t\n",
        ["99"]
    };

    // ── PACK/UNPACK round-trip (8) ────────────────────────────────────

    pack_unpack_int_roundtrip_sum => {
        "program t\ninteger :: src(6) = [1, 2, 3, 4, 5, 6]\nlogical :: mask(6) = [.true., .false., .true., .false., .true., .false.]\ninteger :: tmp(3), dst(6)\ninteger :: fill(6) = [0, 0, 0, 0, 0, 0]\ntmp = pack(src, mask)\ndst = unpack(tmp, mask, fill)\nprint *, sum(dst)\nend program t\n",
        ["9"]
    };
    pack_unpack_int_roundtrip_first_element => {
        "program t\ninteger :: src(5) = [10, 20, 30, 40, 50]\nlogical :: mask(5) = [.true., .false., .true., .false., .true.]\ninteger :: tmp(3), dst(5)\ninteger :: fill(5) = [-1, -1, -1, -1, -1]\ntmp = pack(src, mask)\ndst = unpack(tmp, mask, fill)\nprint *, dst(1)\nprint *, dst(3)\nprint *, dst(5)\nend program t\n",
        ["10", "30", "50"]
    };
    pack_unpack_int_roundtrip_fill_positions => {
        "program t\ninteger :: src(4) = [2, 4, 6, 8]\nlogical :: mask(4) = [.true., .true., .false., .false.]\ninteger :: tmp(2), dst(4)\ninteger :: fill(4) = [100, 100, 100, 100]\ntmp = pack(src, mask)\ndst = unpack(tmp, mask, fill)\nprint *, dst(1)\nprint *, dst(2)\nprint *, dst(3)\nend program t\n",
        ["2", "4", "100"]
    };
    pack_unpack_int_vector_roundtrip => {
        "program t\ninteger :: src(4) = [1, 3, 5, 7]\nlogical :: mask(4) = [.true., .false., .true., .false.]\ninteger :: vec(4) = [0, 0, 0, 0]\ninteger :: tmp(4), dst(4)\ninteger :: fill(4) = [9, 9, 9, 9]\ntmp = pack(src, mask, vec)\ndst = unpack(tmp(1:2), mask, fill)\nprint *, dst(1)\nprint *, dst(3)\nend program t\n",
        ["1", "5"]
    };
    pack_unpack_int_identity_all_true => {
        "program t\ninteger :: src(3) = [7, 8, 9]\nlogical :: mask(3) = [.true., .true., .true.]\ninteger :: tmp(3), dst(3)\ninteger :: fill(3) = [0, 0, 0]\ntmp = pack(src, mask)\ndst = unpack(tmp, mask, fill)\nprint *, dst(1)\nprint *, dst(2)\nprint *, dst(3)\nend program t\n",
        ["7", "8", "9"]
    };
    pack_unpack_int_sparse_restore => {
        "program t\ninteger :: src(8) = [0, 0, 5, 0, 0, 0, 9, 0]\nlogical :: mask(8) = [.false., .false., .true., .false., .false., .false., .true., .false.]\ninteger :: tmp(2), dst(8)\ninteger :: fill(8) = [0, 0, 0, 0, 0, 0, 0, 0]\ntmp = pack(src, mask)\ndst = unpack(tmp, mask, fill)\nprint *, dst(3)\nprint *, dst(7)\nend program t\n",
        ["5", "9"]
    };
    pack_unpack_int_negatives => {
        "program t\ninteger :: src(5) = [-1, 2, -3, 4, -5]\nlogical :: mask(5) = [.true., .false., .true., .false., .true.]\ninteger :: tmp(3), dst(5)\ninteger :: fill(5) = [0, 0, 0, 0, 0]\ntmp = pack(src, mask)\ndst = unpack(tmp, mask, fill)\nprint *, dst(1)\nprint *, dst(3)\nprint *, dst(5)\nend program t\n",
        ["-1", "-3", "-5"]
    };
    pack_unpack_int_count_nonzero => {
        "program t\ninteger :: src(6) = [1, 0, 2, 0, 3, 0]\nlogical :: mask(6) = [.true., .false., .true., .false., .true., .false.]\ninteger :: tmp(3), dst(6)\ninteger :: fill(6) = [0, 0, 0, 0, 0, 0]\ntmp = pack(src, mask)\ndst = unpack(tmp, mask, fill)\nprint *, count(dst > 0)\nend program t\n",
        ["3"]
    };

    // ── Real PACK/UNPACK (8) ──────────────────────────────────────────

    pack_real_alternating_halves => {
        "program t\nreal :: a(4) = [1.5, 2.5, 3.5, 4.5]\nlogical :: mask(4) = [.true., .false., .true., .false.]\nreal :: b(2)\nb = pack(a, mask)\nprint *, int(b(1) + b(2))\nend program t\n",
        ["5"]
    };
    pack_real_all_true_sum => {
        "program t\nreal :: a(3) = [0.5, 1.0, 1.5]\nlogical :: mask(3) = [.true., .true., .true.]\nreal :: b(3)\nb = pack(a, mask)\nprint *, int(sum(b) * 10)\nend program t\n",
        ["30"]
    };
    pack_real_with_vector_pad => {
        "program t\nreal :: a(3) = [2.0, 4.0, 6.0]\nlogical :: mask(3) = [.true., .false., .true.]\nreal :: vec(4) = [0.0, 0.0, 0.0, 0.0]\nreal :: b(4)\nb = pack(a, mask, vec)\nprint *, int(b(1))\nprint *, int(b(2))\nprint *, int(b(4))\nend program t\n",
        ["2", "6", "0"]
    };
    unpack_real_scattered_fill => {
        "program t\nreal :: a(2) = [1.25, 3.75]\nlogical :: mask(4) = [.true., .false., .true., .false.]\nreal :: fill(4) = [0.0, 0.0, 0.0, 0.0]\nreal :: b(4)\nb = unpack(a, mask, fill)\nprint *, int(b(1) * 100)\nprint *, int(b(3) * 100)\nend program t\n",
        ["125", "375"]
    };
    pack_real_negative_values => {
        "program t\nreal :: a(4) = [-1.0, 2.0, -3.0, 4.0]\nlogical :: mask(4) = [.true., .false., .true., .false.]\nreal :: b(2)\nb = pack(a, mask)\nprint *, int(b(1))\nprint *, int(b(2))\nend program t\n",
        ["-1", "-3"]
    };
    pack_unpack_real_roundtrip => {
        "program t\nreal :: src(4) = [1.0, 2.0, 3.0, 4.0]\nlogical :: mask(4) = [.true., .false., .true., .false.]\nreal :: tmp(2), dst(4)\nreal :: fill(4) = [0.0, 0.0, 0.0, 0.0]\ntmp = pack(src, mask)\ndst = unpack(tmp, mask, fill)\nprint *, int(dst(1) + dst(3))\nend program t\n",
        ["4"]
    };
    pack_real_tenths_selection => {
        "program t\nreal :: a(5) = [0.1, 0.2, 0.3, 0.4, 0.5]\nlogical :: mask(5) = [.false., .true., .false., .true., .false.]\nreal :: b(2)\nb = pack(a, mask)\nprint *, int(sum(b) * 10)\nend program t\n",
        ["6"]
    };
    unpack_real_all_false_uses_fill => {
        "program t\nreal :: a(1) = [9.9]\nlogical :: mask(2) = [.false., .false.]\nreal :: fill(2) = [1.1, 2.2]\nreal :: b(2)\nb = unpack(a, mask, fill)\nprint *, int(b(1) * 10)\nprint *, int(b(2) * 10)\nend program t\n",
        ["11", "22"]
    };

    // ── Character PACK/UNPACK (8) ─────────────────────────────────────

    pack_char_two_of_four => {
        "program t\ncharacter(len=1) :: a(4) = ['A', 'B', 'C', 'D']\nlogical :: mask(4) = [.true., .false., .true., .false.]\ncharacter(len=1) :: b(2)\nb = pack(a, mask)\nprint *, b(1)\nprint *, b(2)\nend program t\n",
        ["A", "C"]
    };
    pack_char_all_true_order => {
        "program t\ncharacter(len=1) :: a(3) = ['X', 'Y', 'Z']\nlogical :: mask(3) = [.true., .true., .true.]\ncharacter(len=1) :: b(3)\nb = pack(a, mask)\nprint *, b(1)\nprint *, b(3)\nend program t\n",
        ["X", "Z"]
    };
    pack_char_with_vector_pad => {
        "program t\ncharacter(len=1) :: a(2) = ['P', 'Q']\nlogical :: mask(2) = [.true., .true.]\ncharacter(len=1) :: vec(4) = ['-', '-', '-', '-']\ncharacter(len=1) :: b(4)\nb = pack(a, mask, vec)\nprint *, b(1)\nprint *, b(2)\nprint *, b(3)\nend program t\n",
        ["P", "Q", "-"]
    };
    unpack_char_alternating_fill => {
        "program t\ncharacter(len=1) :: a(2) = ['M', 'N']\nlogical :: mask(4) = [.true., .false., .true., .false.]\ncharacter(len=1) :: fill(4) = ['.', '.', '.', '.']\ncharacter(len=1) :: b(4)\nb = unpack(a, mask, fill)\nprint *, b(1)\nprint *, b(2)\nprint *, b(3)\nend program t\n",
        ["M", ".", "N"]
    };
    pack_char_first_last => {
        "program t\ncharacter(len=1) :: a(5) = ['1', '2', '3', '4', '5']\nlogical :: mask(5) = [.true., .false., .false., .false., .true.]\ncharacter(len=1) :: b(2)\nb = pack(a, mask)\nprint *, b(1)\nprint *, b(2)\nend program t\n",
        ["1", "5"]
    };
    unpack_char_all_false => {
        "program t\ncharacter(len=1) :: a(1) = ['Z']\nlogical :: mask(3) = [.false., .false., .false.]\ncharacter(len=1) :: fill(3) = ['a', 'b', 'c']\ncharacter(len=1) :: b(3)\nb = unpack(a, mask, fill)\nprint *, b(1)\nprint *, b(2)\nend program t\n",
        ["a", "b"]
    };
    pack_unpack_char_roundtrip => {
        "program t\ncharacter(len=1) :: src(4) = ['W', 'X', 'Y', 'Z']\nlogical :: mask(4) = [.true., .false., .true., .false.]\ncharacter(len=1) :: tmp(2), dst(4)\ncharacter(len=1) :: fill(4) = ['?', '?', '?', '?']\ntmp = pack(src, mask)\ndst = unpack(tmp, mask, fill)\nprint *, dst(1)\nprint *, dst(3)\nend program t\n",
        ["W", "Y"]
    };
    pack_char_vowels_from_word => {
        "program t\ncharacter(len=1) :: a(5) = ['H', 'E', 'L', 'L', 'O']\nlogical :: mask(5) = [.false., .true., .false., .false., .true.]\ncharacter(len=1) :: b(2)\nb = pack(a, mask)\nprint *, b(1)\nprint *, b(2)\nend program t\n",
        ["E", "O"]
    };

    // ── 2D PACK/UNPACK (5) ────────────────────────────────────────────

    pack_2d_by_column_major_mask => {
        "program t\ninteger :: a(2,3) = reshape([1, 4, 2, 5, 3, 6], [2, 3])\nlogical :: mask(2,3) = reshape([.true., .false., .true., .false., .true., .false.], [2, 3])\ninteger :: b(3)\nb = pack(a, mask)\nprint *, b(1)\nprint *, b(2)\nprint *, b(3)\nend program t\n",
        ["1", "2", "3"]
    };
    pack_2d_row_mask_three_selected => {
        "program t\ninteger :: a(3,2) = reshape([1, 2, 3, 4, 5, 6], [3, 2])\nlogical :: mask(3,2) = reshape([.true., .true., .false., .false., .true., .false.], [3, 2])\ninteger :: b(3)\nb = pack(a, mask)\nprint *, sum(b)\nend program t\n",
        ["9"]
    };
    unpack_2d_restore_matrix => {
        "program t\ninteger :: a(3) = [10, 20, 30]\nlogical :: mask(2,2) = reshape([.true., .false., .true., .false.], [2, 2])\ninteger :: fill(2,2) = 0\ninteger :: b(2,2)\nb = unpack(a, mask, fill)\nprint *, b(1,1)\nprint *, b(1,2)\nprint *, b(2,1)\nend program t\n",
        ["10", "0", "20"]
    };
    pack_2d_with_vector_pad => {
        "program t\ninteger :: a(2,2) = reshape([1, 2, 3, 4], [2, 2])\nlogical :: mask(2,2) = reshape([.true., .false., .true., .false.], [2, 2])\ninteger :: vec(3) = [0, 0, 0]\ninteger :: b(3)\nb = pack(a, mask, vec)\nprint *, b(1)\nprint *, b(2)\nprint *, b(3)\nend program t\n",
        ["1", "3", "0"]
    };
    pack_unpack_2d_roundtrip_corner => {
        "program t\ninteger :: src(2,2) = reshape([5, 6, 7, 8], [2, 2])\nlogical :: mask(2,2) = reshape([.true., .false., .false., .true.], [2, 2])\ninteger :: tmp(2), dst(2,2)\ninteger :: fill(2,2) = 0\ntmp = pack(src, mask)\ndst = unpack(tmp, mask, fill)\nprint *, dst(1,1)\nprint *, dst(2,2)\nend program t\n",
        ["5", "8"]
    };
}
