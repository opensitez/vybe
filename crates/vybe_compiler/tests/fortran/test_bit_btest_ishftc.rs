//! BTEST and ISHFTC bit intrinsics, negative-integer bit positions, and
//! combinations with IBSET/IBCLR/IAND/IEOR/ISHFT not covered in `test_bits_f2008.rs`
//! (popcount/leadz/trailz) or `test_intrinsics_extended.rs` (basic iand/ior/ieor).

use super::helpers::compile_ok;

fortran_cases! {
    // ── Negative-integer bit positions (runtime via iand/ibset/ibclr) ──

    iand_negative_one_bit_zero => {
        "program t\nprint *, iand(-1, ishft(1, 0))\nend program t\n",
        ["1"]
    };

    iand_negative_one_sign_bit => {
        "program t\nprint *, iand(-1, ishft(1, 31))\nend program t\n",
        ["-2147483648"]
    };

    ibset_negative_two_bit_zero => {
        "program t\nprint *, ibset(-2, 0)\nend program t\n",
        ["-1"]
    };

    ibclr_negative_one_bit_zero => {
        "program t\nprint *, ibclr(-1, 0)\nend program t\n",
        ["-2"]
    };

    ieor_negative_one_with_one => {
        "program t\nprint *, ieor(-1, 1)\nend program t\n",
        ["-2"]
    };

    ishft_negative_two_right_one => {
        "program t\nprint *, ishft(-2, -1)\nend program t\n",
        ["-1"]
    };

    ibset_then_iand_same_bit_position => {
        "program t\ninteger :: x\nx = ibset(0, 5)\nprint *, x\nprint *, iand(x, ishft(1, 5))\nend program t\n",
        ["32", "32"]
    };

    ieor_clears_set_bit_via_mask => {
        "program t\nprint *, ieor(255, ishft(1, 4))\nend program t\n",
        ["239"]
    };

    iand_negative_one_with_byte_mask => {
        "program t\nprint *, iand(-1, 255)\nend program t\n",
        ["255"]
    };

    not_negative_two => {
        "program t\nprint *, not(-2)\nend program t\n",
        ["1"]
    };
}

// ── BTEST at bit positions (compile-only until intrinsic is lowered) ──

#[test]
fn btest_bit_three_of_eight() {
    compile_ok("program t\n  print *, btest(8, 3)\nend program t\n");
}

#[test]
fn btest_bit_zero_of_eight_is_clear() {
    compile_ok("program t\n  print *, btest(8, 0)\nend program t\n");
}

#[test]
fn btest_negative_one_bit_zero() {
    compile_ok("program t\n  print *, btest(-1, 0)\nend program t\n");
}

#[test]
fn btest_negative_two_bit_zero() {
    compile_ok("program t\n  print *, btest(-2, 0)\nend program t\n");
}

#[test]
fn btest_negative_one_sign_bit() {
    compile_ok("program t\n  print *, btest(-1, 31)\nend program t\n");
}

#[test]
fn btest_scan_lower_nibble() {
    compile_ok(
        r#"
program t
    integer :: i, x = 10
    do i = 0, 3
        print *, btest(x, i)
    end do
end program t
"#,
    );
}

#[test]
fn btest_after_ibset_same_position() {
    compile_ok(
        r#"
program t
    integer :: x
    x = ibset(0, 7)
    print *, btest(x, 7)
    print *, btest(x, 6)
end program t
"#,
    );
}

#[test]
fn btest_with_ieor_in_expression() {
    compile_ok(
        r#"
program t
    integer :: x = 42
    print *, btest(ieor(x, ishft(1, 1)), 1)
end program t
"#,
    );
}

// ── ISHFTC circular shift within size field (compile-only) ─────────

#[test]
fn ishftc_rotate_left_within_nibble() {
    compile_ok("program t\n  print *, ishftc(12, 2, 4)\nend program t\n");
}

#[test]
fn ishftc_rotate_right_within_nibble() {
    compile_ok("program t\n  print *, ishftc(12, -2, 4)\nend program t\n");
}

#[test]
fn ishftc_single_bit_left_in_four_bits() {
    compile_ok("program t\n  print *, ishftc(1, 1, 4)\nend program t\n");
}

#[test]
fn ishftc_single_bit_right_in_four_bits() {
    compile_ok("program t\n  print *, ishftc(8, -1, 4)\nend program t\n");
}

#[test]
fn ishftc_zero_shift_preserves_field() {
    compile_ok("program t\n  print *, ishftc(170, 0, 8)\nend program t\n");
}

#[test]
fn ishftc_full_byte_rotate() {
    compile_ok("program t\n  print *, ishftc(170, 1, 8)\nend program t\n");
}

#[test]
fn ishftc_with_variable_size_operand() {
    compile_ok(
        r#"
program t
    integer :: n = 4
    print *, ishftc(6, 1, n)
end program t
"#,
    );
}

#[test]
fn ishftc_assign_then_btest_result_bit() {
    compile_ok(
        r#"
program t
    integer :: x
    x = ishftc(5, 2, 4)
    print *, x
    print *, btest(x, 2)
end program t
"#,
    );
}
