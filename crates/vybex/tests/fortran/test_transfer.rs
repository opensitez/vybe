use super::helpers::compile_ok;

// ── TRANSFER — scalar to scalar ───────────────────────────────

#[test] fn transfer_int_to_real() {
    compile_ok(r#"
program test
    integer :: i = 0
    real :: r
    r = transfer(i, 0.0)
    print *, 'ok'
end program test
"#);
}

#[test] fn transfer_real_to_int() {
    compile_ok(r#"
program test
    real :: x = 0.0
    integer :: i
    i = transfer(x, 0)
    print *, 'ok'
end program test
"#);
}

#[test] fn transfer_int_to_int_same() {
    compile_ok(r#"
program test
    integer :: a = 42, b
    b = transfer(a, 0)
    print *, b
end program test
"#);
}

#[test] fn transfer_double_to_two_ints() {
    compile_ok(r#"
program test
    real(kind=8) :: d = 0.0d0
    integer :: parts(2)
    parts = transfer(d, parts)
    print *, 'ok'
end program test
"#);
}

#[test] fn transfer_complex_to_real_pair() {
    compile_ok(r#"
program test
    complex :: c = (1.0, 2.0)
    real :: pair(2)
    pair = transfer(c, pair)
    print *, 'ok'
end program test
"#);
}

#[test] fn transfer_real_pair_to_complex() {
    compile_ok(r#"
program test
    real :: pair(2) = [1.0, 2.0]
    complex :: c
    c = transfer(pair, c)
    print *, 'ok'
end program test
"#);
}

// ── TRANSFER with SIZE parameter ──────────────────────────────

#[test] fn transfer_with_size() {
    compile_ok(r#"
program test
    integer :: a(4) = [1, 2, 3, 4]
    integer(kind=1) :: bytes(16)
    bytes = transfer(a, bytes)
    print *, 'ok'
end program test
"#);
}

#[test] fn transfer_size_truncate() {
    compile_ok(r#"
program test
    integer :: a(4) = [1, 2, 3, 4]
    integer :: b(2)
    b = transfer(a, b, 2)
    print *, b(1)
end program test
"#);
}

#[test] fn transfer_size_expand() {
    compile_ok(r#"
program test
    integer :: a = 42
    integer :: b(4)
    b = transfer(a, b, 4)
    print *, 'ok'
end program test
"#);
}

// ── TRANSFER — character ──────────────────────────────────────

#[test] fn transfer_char_to_int() {
    compile_ok(r#"
program test
    character(len=4) :: s = 'ABCD'
    integer :: n
    n = transfer(s, 0)
    print *, 'ok'
end program test
"#);
}

#[test] fn transfer_int_to_char() {
    compile_ok(r#"
program test
    integer :: n = 1195853639
    character(len=4) :: s
    s = transfer(n, '    ')
    print *, 'ok'
end program test
"#);
}

// ── TRANSFER — scalar to/from array ───────────────────────────

#[test] fn transfer_scalar_to_array() {
    compile_ok(r#"
program test
    integer(kind=8) :: big = 1000000000000_8
    integer(kind=4) :: parts(2)
    parts = transfer(big, parts)
    print *, 'ok'
end program test
"#);
}

#[test] fn transfer_array_to_scalar() {
    compile_ok(r#"
program test
    integer(kind=4) :: parts(2) = [0, 0]
    integer(kind=8) :: big
    big = transfer(parts, 0_8)
    print *, 'ok'
end program test
"#);
}

// ── TRANSFER — derived types ──────────────────────────────────

#[test] fn transfer_derived_type_to_array() {
    compile_ok(r#"
program test
    type :: Point
        real :: x, y
    end type Point
    type(Point) :: p
    real :: coords(2)
    p%x = 1.0; p%y = 2.0
    coords = transfer(p, coords)
    print *, 'ok'
end program test
"#);
}

#[test] fn transfer_array_to_derived_type() {
    compile_ok(r#"
program test
    type :: Point
        real :: x, y
    end type Point
    real :: coords(2) = [3.0, 4.0]
    type(Point) :: p
    p = transfer(coords, p)
    print *, 'ok'
end program test
"#);
}

#[test] fn transfer_sequence_type() {
    compile_ok(r#"
program test
    type, sequence :: RGB
        integer(kind=1) :: r, g, b, a
    end type RGB
    type(RGB) :: color
    integer :: packed
    color%r = 255_1; color%g = 0_1; color%b = 128_1; color%a = 255_1
    packed = transfer(color, 0)
    print *, 'ok'
end program test
"#);
}

// ── TRANSFER in functions/subroutines ────────────────────────

#[test] fn transfer_in_function() {
    compile_ok(r#"
program test
    print *, real_bits(1.0)
contains
    function real_bits(x) result(n)
        real, intent(in) :: x
        integer :: n
        n = transfer(x, 0)
    end function real_bits
end program test
"#);
}

#[test] fn transfer_in_subroutine() {
    compile_ok(r#"
program test
    real :: x = 3.14
    integer :: n
    call get_bits(x, n)
    print *, 'ok'
contains
    subroutine get_bits(x, n)
        real, intent(in) :: x
        integer, intent(out) :: n
        n = transfer(x, 0)
    end subroutine get_bits
end program test
"#);
}

// ── TRANSFER endianness / byte-level ─────────────────────────

#[test] fn transfer_byte_array_roundtrip() {
    compile_ok(r#"
program test
    integer :: original = 305419896
    integer(kind=1) :: bytes(4)
    integer :: recovered
    bytes = transfer(original, bytes)
    recovered = transfer(bytes, 0)
    print *, original == recovered
end program test
"#);
}

#[test] fn transfer_kind8_bytes() {
    compile_ok(r#"
program test
    integer(kind=8) :: n = 1_8
    integer(kind=1) :: bytes(8)
    bytes = transfer(n, bytes)
    print *, 'ok'
end program test
"#);
}
