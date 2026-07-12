use super::helpers::compile_ok;

// ── POPCOUNT ──────────────────────────────────────────────────

#[test]
fn popcount_zero() {
    compile_ok("program t\n  print *, popcount(0)\nend program t\n");
}

#[test]
fn popcount_one() {
    compile_ok("program t\n  print *, popcount(1)\nend program t\n");
}

#[test]
fn popcount_255() {
    compile_ok("program t\n  print *, popcount(255)\nend program t\n");
}

#[test]
fn popcount_int8() {
    compile_ok(
        "program t\n  integer(kind=8) :: x = 1152921504606846975_8\n  print *, popcount(x)\nend program t\n",
    );
}

#[test]
fn popcount_negative() {
    compile_ok("program t\n  integer :: x = -1\n  print *, popcount(x)\nend program t\n");
}

#[test]
fn popcount_in_array() {
    compile_ok(
        r#"
program test
    integer :: a(4) = [0, 1, 3, 7]
    integer :: i
    do i = 1, 4
        print *, popcount(a(i))
    end do
end program test
"#,
    );
}

// ── LEADZ ─────────────────────────────────────────────────────

#[test]
fn leadz_one() {
    compile_ok("program t\n  print *, leadz(1)\nend program t\n");
}

#[test]
fn leadz_max_int() {
    compile_ok("program t\n  print *, leadz(huge(0))\nend program t\n");
}

#[test]
fn leadz_zero() {
    compile_ok("program t\n  print *, leadz(0)\nend program t\n");
}

#[test]
fn leadz_kind8() {
    compile_ok("program t\n  integer(kind=8) :: x = 1_8\n  print *, leadz(x)\nend program t\n");
}

#[test]
fn leadz_power_of_two() {
    compile_ok(
        r#"
program test
    integer :: i
    do i = 0, 7
        print *, leadz(2**i)
    end do
end program test
"#,
    );
}

// ── TRAILZ ────────────────────────────────────────────────────

#[test]
fn trailz_one() {
    compile_ok("program t\n  print *, trailz(1)\nend program t\n");
}

#[test]
fn trailz_two() {
    compile_ok("program t\n  print *, trailz(2)\nend program t\n");
}

#[test]
fn trailz_eight() {
    compile_ok("program t\n  print *, trailz(8)\nend program t\n");
}

#[test]
fn trailz_zero() {
    compile_ok("program t\n  print *, trailz(0)\nend program t\n");
}

#[test]
fn trailz_and_leadz_together() {
    compile_ok(
        r#"
program test
    integer :: x = 16
    print *, leadz(x)
    print *, trailz(x)
    print *, popcount(x)
end program test
"#,
    );
}

// ── DSHIFTL / DSHIFTR ─────────────────────────────────────────

#[test]
fn dshiftl_basic() {
    compile_ok("program t\n  print *, dshiftl(0, 1, 1)\nend program t\n");
}

#[test]
fn dshiftl_carries_bit() {
    compile_ok(
        r#"
program test
    integer :: hi = int(z'80000000')
    integer :: lo = 0
    print *, dshiftl(hi, lo, 1)
end program test
"#,
    );
}

#[test]
fn dshiftr_basic() {
    compile_ok("program t\n  print *, dshiftr(1, 0, 1)\nend program t\n");
}

#[test]
fn dshiftr_carries_bit() {
    compile_ok(
        r#"
program test
    integer :: hi = 1
    integer :: lo = 0
    print *, dshiftr(hi, lo, 1)
end program test
"#,
    );
}

#[test]
fn dshiftl_zero_shift() {
    compile_ok("program t\n  print *, dshiftl(42, 0, 0)\nend program t\n");
}

// ── MASKL / MASKR ─────────────────────────────────────────────

#[test]
fn maskl_basic() {
    compile_ok("program t\n  print *, maskl(4)\nend program t\n");
}

#[test]
fn maskl_zero() {
    compile_ok("program t\n  print *, maskl(0)\nend program t\n");
}

#[test]
fn maskl_full() {
    compile_ok("program t\n  print *, maskl(bit_size(0))\nend program t\n");
}

#[test]
fn maskr_basic() {
    compile_ok("program t\n  print *, maskr(4)\nend program t\n");
}

#[test]
fn maskr_zero() {
    compile_ok("program t\n  print *, maskr(0)\nend program t\n");
}

#[test]
fn maskr_full() {
    compile_ok("program t\n  print *, maskr(bit_size(0))\nend program t\n");
}

#[test]
fn maskl_maskr_complement() {
    compile_ok(
        r#"
program test
    integer :: n = 4
    integer :: l, r
    l = maskl(n)
    r = maskr(bit_size(0) - n)
    print *, l == r
end program test
"#,
    );
}

// ── MERGE_BITS ────────────────────────────────────────────────

#[test]
fn merge_bits_basic() {
    compile_ok(
        "program t\n  print *, merge_bits(int(z'FF00'), int(z'00FF'), int(z'F0F0'))\nend program t\n",
    );
}

#[test]
fn merge_bits_all_from_i() {
    compile_ok("program t\n  print *, merge_bits(255, 0, 255)\nend program t\n");
}

#[test]
fn merge_bits_all_from_j() {
    compile_ok("program t\n  print *, merge_bits(255, 0, 0)\nend program t\n");
}

#[test]
fn merge_bits_alternating() {
    compile_ok(
        r#"
program test
    integer :: result
    result = merge_bits(int(z'AAAA'), int(z'5555'), int(z'FF00'))
    print *, result
end program test
"#,
    );
}

// ── BGE / BGT / BLE / BLT ─────────────────────────────────────

#[test]
fn bge_equal() {
    compile_ok("program t\n  print *, bge(5, 5)\nend program t\n");
}

#[test]
fn bge_greater() {
    compile_ok("program t\n  print *, bge(6, 5)\nend program t\n");
}

#[test]
fn bge_less() {
    compile_ok("program t\n  print *, bge(4, 5)\nend program t\n");
}

#[test]
fn bgt_equal() {
    compile_ok("program t\n  print *, bgt(5, 5)\nend program t\n");
}

#[test]
fn bgt_greater() {
    compile_ok("program t\n  print *, bgt(6, 5)\nend program t\n");
}

#[test]
fn ble_equal() {
    compile_ok("program t\n  print *, ble(5, 5)\nend program t\n");
}

#[test]
fn ble_less() {
    compile_ok("program t\n  print *, ble(4, 5)\nend program t\n");
}

#[test]
fn blt_less() {
    compile_ok("program t\n  print *, blt(4, 5)\nend program t\n");
}

#[test]
fn blt_equal() {
    compile_ok("program t\n  print *, blt(5, 5)\nend program t\n");
}

#[test]
fn bitwise_compare_negative() {
    compile_ok(
        r#"
program test
    integer :: a = -1, b = 1
    print *, bgt(a, b)
    print *, blt(b, a)
end program test
"#,
    );
}

// ── PARITY ────────────────────────────────────────────────────

#[test]
fn parity_single_true() {
    compile_ok("program t\n  logical :: a(1) = [.true.]\n  print *, parity(a)\nend program t\n");
}

#[test]
fn parity_two_true() {
    compile_ok(
        "program t\n  logical :: a(2) = [.true., .true.]\n  print *, parity(a)\nend program t\n",
    );
}

#[test]
fn parity_three_true() {
    compile_ok(
        "program t\n  logical :: a(3) = [.true., .true., .true.]\n  print *, parity(a)\nend program t\n",
    );
}

#[test]
fn parity_mixed() {
    compile_ok(
        r#"
program test
    logical :: a(4) = [.true., .false., .true., .false.]
    print *, parity(a)
end program test
"#,
    );
}

#[test]
fn parity_empty_is_false() {
    compile_ok(
        r#"
program test
    logical :: a(0)
    print *, parity(a)
end program test
"#,
    );
}

#[test]
fn parity_with_dim() {
    compile_ok(
        r#"
program test
    logical :: m(2,3) = reshape([.true.,.false.,.true.,.false.,.true.,.false.],[2,3])
    logical :: row_parity(3)
    row_parity = parity(m, dim=1)
    print *, row_parity(1)
end program test
"#,
    );
}
