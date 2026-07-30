use super::helpers::run_prints;

// ── TRANSFER — scalar to scalar ───────────────────────────────

#[test]
fn transfer_int_to_real() {
    let out = run_prints(
        r#"
program test
    integer :: i = 0
    real :: r
    r = transfer(i, 0.0)
    print *, r
end program test
"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn transfer_real_to_int() {
    let out = run_prints(
        r#"
program test
    real :: x = 0.0
    integer :: i
    i = transfer(x, 0)
    print *, i
end program test
"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn transfer_int_to_int_same() {
    let out = run_prints(
        r#"
program test
    integer :: a = 42, b
    b = transfer(a, 0)
    print *, b
end program test
"#,
    );
    assert_eq!(out, vec!["42"]);
}

#[test]
fn transfer_double_to_two_ints() {
    let out = run_prints(
        r#"
program test
    real(kind=8) :: d = 0.0d0
    integer :: parts(2)
    parts = transfer(d, parts)
    print *, parts(1)
    print *, parts(2)
end program test
"#,
    );
    assert_eq!(out, vec!["0", "0"]);
}

#[test]
fn transfer_complex_to_real_pair() {
    let out = run_prints(
        r#"
program test
    complex :: c = (1.0, 2.0)
    real :: pair(2)
    pair = transfer(c, pair)
    print *, pair(1)
    print *, pair(2)
end program test
"#,
    );
    assert_eq!(out, vec!["1", "2"]);
}

#[test]
fn transfer_real_pair_to_complex() {
    let out = run_prints(
        r#"
program test
    real :: pair(2) = [1.0, 2.0]
    complex :: c
    c = transfer(pair, c)
    print *, real(c)
    print *, imag(c)
end program test
"#,
    );
    assert_eq!(out, vec!["1", "2"]);
}

// ── TRANSFER with SIZE parameter ──────────────────────────────

#[test]
fn transfer_with_size() {
    let out = run_prints(
        r#"
program test
    integer :: a(4) = [0, 0, 0, 0]
    integer(kind=1) :: bytes(16)
    bytes = transfer(a, bytes)
    print *, bytes(1)
    print *, bytes(2)
    print *, bytes(3)
end program test
"#,
    );
    assert_eq!(out, vec!["0", "0", "0"]);
}

#[test]
fn transfer_size_truncate() {
    let out = run_prints(
        r#"
program test
    integer :: a(4) = [1, 2, 3, 4]
    integer :: b(2)
    b = transfer(a, b, 2)
    print *, b(1)
end program test
"#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn transfer_size_expand() {
    let out = run_prints(
        r#"
program test
    integer :: a = 42
    integer :: b(4)
    b = transfer(a, b, 4)
    print *, b(1)
    print *, b(2)
    print *, b(3)
end program test
"#,
    );
    assert_eq!(out, vec!["42", "0", "0"]);
}

// ── TRANSFER — character ──────────────────────────────────────

#[test]
fn transfer_char_to_int() {
    let out = run_prints(
        r#"
program test
    character(len=4) :: s
    integer :: n
    s = transfer(0, s)
    n = transfer(s, 0)
    print *, n
end program test
"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn transfer_int_to_char() {
    let out = run_prints(
        r#"
program test
    integer :: n = 1195853639
    character(len=4) :: s
    s = transfer(n, '    ')
    print *, len(s)
end program test
"#,
    );
    assert_eq!(out, vec!["4"]);
}

// ── TRANSFER — scalar to/from array ───────────────────────────

#[test]
fn transfer_scalar_to_array() {
    let out = run_prints(
        r#"
program test
    integer(kind=8) :: big = 0_8
    integer(kind=4) :: parts(2)
    parts = transfer(big, parts)
    print *, parts(1)
    print *, parts(2)
end program test
"#,
    );
    assert_eq!(out, vec!["0", "0"]);
}

#[test]
fn transfer_array_to_scalar() {
    let out = run_prints(
        r#"
program test
    integer(kind=4) :: parts(2) = [0, 0]
    integer(kind=8) :: big
    big = transfer(parts, 0_8)
    print *, big == 0
end program test
"#,
    );
    assert_eq!(out, vec!["1"]);
}

// ── TRANSFER — derived types ──────────────────────────────────

#[test]
fn transfer_derived_type_to_array() {
    let out = run_prints(
        r#"
program test
    type :: Point
        real :: x, y
    end type Point
    type(Point) :: p
    real :: coords(2)
    p%x = 1.0; p%y = 2.0
    coords = transfer(p, coords)
    print *, coords(1)
    print *, coords(2)
end program test
"#,
    );
    assert_eq!(out, vec!["1", "2"]);
}

#[test]
fn transfer_array_to_derived_type() {
    let out = run_prints(
        r#"
program test
    type :: Point
        real :: x, y
    end type Point
    real :: coords(2) = [3.0, 4.0]
    type(Point) :: p
    p = transfer(coords, p)
    print *, p%x
    print *, p%y
end program test
"#,
    );
    assert_eq!(out, vec!["3", "4"]);
}

#[test]
fn transfer_sequence_type() {
    let out = run_prints(
        r#"
program test
    type, sequence :: RGB
        integer(kind=1) :: r, g, b, a
    end type RGB
    type(RGB) :: color
    integer :: packed
    color%r = 255_1; color%g = 0_1; color%b = 128_1; color%a = 255_1
    packed = transfer(color, 0)
    print *, color%r
    print *, color%g
    print *, color%b
    print *, color%a
end program test
"#,
    );
    assert_eq!(out, vec!["255", "0", "128", "255"]);
}

// ── TRANSFER in functions/subroutines ────────────────────────

#[test]
fn transfer_in_function() {
    let out = run_prints(
        r#"
program test
    print *, real_bits_roundtrip(1.0)
contains
    logical function real_bits_roundtrip(x)
        real, intent(in) :: x
        integer :: n
        n = transfer(x, 0)
        real_bits_roundtrip = (transfer(n, 0.0) == x)
    end function real_bits_roundtrip
end program test
"#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn transfer_in_subroutine() {
    let out = run_prints(
        r#"
program test
    real :: x = 3.14
    logical :: same
    call get_bits(x, same)
    print *, same
contains
    subroutine get_bits(x, same)
        real, intent(in) :: x
        logical, intent(out) :: same
        integer :: n
        n = transfer(x, 0)
        same = transfer(n, 0.0) == x
    end subroutine get_bits
end program test
"#,
    );
    assert_eq!(out, vec!["1"]);
}

// ── TRANSFER endianness / byte-level ─────────────────────────

#[test]
fn transfer_byte_array_roundtrip() {
    let out = run_prints(
        r#"
program test
    integer :: original = 305419896
    integer(kind=1) :: bytes(4)
    integer :: recovered
    bytes = transfer(original, bytes)
    recovered = transfer(bytes, 0)
    print *, original == recovered
end program test
"#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn transfer_kind8_bytes() {
    let out = run_prints(
        r#"
program test
    integer(kind=8) :: n = 0_8
    integer(kind=1) :: bytes(8)
    bytes = transfer(n, bytes)
    print *, bytes(1)
end program test
"#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn transfer_int_roundtrip_same_type() {
    let out = run_prints(
        r#"
program test
    integer :: original
    integer :: copy
    original = -17
    copy = transfer(original, 0)
    print *, copy
    print *, original == copy
end program test
"#,
    );
    assert_eq!(out, vec!["-17", "1"]);
}

#[test]
fn transfer_int_real_roundtrip_runtime() {
    let out = run_prints(
        r#"
program test
    integer :: original
    integer :: recovered
    real :: bits
    original = 42
    bits = transfer(original, 0.0)
    recovered = transfer(bits, 0)
    print *, recovered
    print *, original == recovered
end program test
"#,
    );
    assert_eq!(out, vec!["42", "1"]);
}

#[test]
fn transfer_array_roundtrip_and_slice() {
    let out = run_prints(
        r#"
program test
    integer :: source(2) = [11, 22]
    integer :: target(2)
    target = transfer(source, target)
    print *, target(1)
    print *, target(2)
end program test
"#,
    );
    assert_eq!(out, vec!["11", "22"]);
}

#[test]
fn transfer_size_truncates_to_requested_length() {
    let out = run_prints(
        r#"
program test
    integer :: source(4) = [10, 20, 30, 40]
    integer :: target(2)
    target = transfer(source, target, 2)
    print *, target(1)
    print *, target(2)
end program test
"#,
    );
    assert_eq!(out, vec!["10", "20"]);
}

#[test]
fn transfer_size_expands_with_zero_padding() {
    let out = run_prints(
        r#"
program test
    integer :: source
    integer :: target(4)
    source = 99
    target = transfer(source, target, 4)
    print *, target(1)
    print *, target(2)
    print *, target(3)
    print *, target(4)
end program test
"#,
    );
    assert_eq!(out, vec!["99", "0", "0", "0"]);
}

#[test]
fn transfer_logical_values_to_integer() {
    let out = run_prints(
        r#"
program test
    logical :: l1, l2
    integer :: bits(2)
    l1 = .true.
    l2 = .false.
    bits = transfer([l1, l2], bits)
    print *, bits(1)
    print *, bits(2)
end program test
"#,
    );
    assert_eq!(out, vec!["1", "0"]);
}

#[test]
fn transfer_sequence_type_runtime_roundtrip() {
    let out = run_prints(
        r#"
program test
    type, sequence :: RGB
        integer(kind=1) :: r, g, b, a
    end type RGB
    type(RGB) :: in_colour
    type(RGB) :: out_colour
    integer(kind=1) :: bytes(4)

    in_colour%r = 1_1
    in_colour%g = 2_1
    in_colour%b = 3_1
    in_colour%a = 4_1
    bytes = transfer(in_colour, bytes)
    out_colour = transfer(bytes, out_colour)

    print *, in_colour%r == out_colour%r
    print *, in_colour%g == out_colour%g
    print *, in_colour%b == out_colour%b
    print *, in_colour%a == out_colour%a
end program test
"#,
    );
    assert_eq!(out, vec!["1", "1", "1", "1"]);
}
