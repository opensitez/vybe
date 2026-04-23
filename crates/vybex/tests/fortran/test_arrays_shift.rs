use super::helpers::compile_ok;

// ── CSHIFT — 1D ───────────────────────────────────────────────

#[test] fn cshift_1d_left_one() {
    compile_ok(r#"
program test
    integer :: a(5) = [1, 2, 3, 4, 5]
    integer :: b(5)
    b = cshift(a, 1)
    print *, b(1)
    print *, b(5)
end program test
"#);
}

#[test] fn cshift_1d_left_two() {
    compile_ok(r#"
program test
    integer :: a(5) = [1, 2, 3, 4, 5]
    integer :: b(5)
    b = cshift(a, 2)
    print *, b(1)
end program test
"#);
}

#[test] fn cshift_1d_right_one() {
    compile_ok(r#"
program test
    integer :: a(5) = [1, 2, 3, 4, 5]
    integer :: b(5)
    b = cshift(a, -1)
    print *, b(1)
    print *, b(5)
end program test
"#);
}

#[test] fn cshift_1d_right_two() {
    compile_ok(r#"
program test
    integer :: a(6) = [1, 2, 3, 4, 5, 6]
    integer :: b(6)
    b = cshift(a, -2)
    print *, b(1)
    print *, b(2)
end program test
"#);
}

#[test] fn cshift_1d_full_rotation() {
    compile_ok(r#"
program test
    integer :: a(4) = [1, 2, 3, 4]
    integer :: b(4)
    b = cshift(a, 4)
    print *, b(1)
end program test
"#);
}

#[test] fn cshift_1d_zero() {
    compile_ok(r#"
program test
    integer :: a(5) = [1, 2, 3, 4, 5]
    integer :: b(5)
    b = cshift(a, 0)
    print *, b(1)
    print *, b(5)
end program test
"#);
}

#[test] fn cshift_1d_real() {
    compile_ok(r#"
program test
    real :: a(4) = [1.0, 2.0, 3.0, 4.0]
    real :: b(4)
    b = cshift(a, 1)
    print *, b(1)
end program test
"#);
}

#[test] fn cshift_1d_logical() {
    compile_ok(r#"
program test
    logical :: a(4) = [.true., .false., .true., .false.]
    logical :: b(4)
    b = cshift(a, 1)
    print *, b(1)
end program test
"#);
}

// ── CSHIFT — 2D with DIM ──────────────────────────────────────

#[test] fn cshift_2d_dim1() {
    compile_ok(r#"
program test
    integer :: m(3,4) = reshape([(i, i=1,12)],[3,4])
    integer :: n(3,4)
    n = cshift(m, 1, dim=1)
    print *, n(1,1)
end program test
"#);
}

#[test] fn cshift_2d_dim2() {
    compile_ok(r#"
program test
    integer :: m(3,4) = reshape([(i, i=1,12)],[3,4])
    integer :: n(3,4)
    n = cshift(m, 1, dim=2)
    print *, n(1,1)
end program test
"#);
}

#[test] fn cshift_2d_negative_dim2() {
    compile_ok(r#"
program test
    integer :: m(2,4) = reshape([1,2,3,4,5,6,7,8],[2,4])
    integer :: n(2,4)
    n = cshift(m, -1, dim=2)
    print *, n(1,1)
end program test
"#);
}

#[test] fn cshift_2d_with_shift_array() {
    compile_ok(r#"
program test
    integer :: m(3,4) = reshape([(i, i=1,12)],[3,4])
    integer :: shifts(4) = [0, 1, 2, 1]
    integer :: n(3,4)
    n = cshift(m, shifts, dim=1)
    print *, n(1,1)
end program test
"#);
}

// ── EOSHIFT — 1D ──────────────────────────────────────────────

#[test] fn eoshift_1d_left_one() {
    compile_ok(r#"
program test
    integer :: a(5) = [1, 2, 3, 4, 5]
    integer :: b(5)
    b = eoshift(a, 1)
    print *, b(1)
    print *, b(5)
end program test
"#);
}

#[test] fn eoshift_1d_left_two() {
    compile_ok(r#"
program test
    integer :: a(5) = [1, 2, 3, 4, 5]
    integer :: b(5)
    b = eoshift(a, 2)
    print *, b(1)
    print *, b(4)
    print *, b(5)
end program test
"#);
}

#[test] fn eoshift_1d_right_one() {
    compile_ok(r#"
program test
    integer :: a(5) = [1, 2, 3, 4, 5]
    integer :: b(5)
    b = eoshift(a, -1)
    print *, b(1)
    print *, b(5)
end program test
"#);
}

#[test] fn eoshift_1d_with_boundary() {
    compile_ok(r#"
program test
    integer :: a(5) = [1, 2, 3, 4, 5]
    integer :: b(5)
    b = eoshift(a, 2, boundary=-1)
    print *, b(4)
    print *, b(5)
end program test
"#);
}

#[test] fn eoshift_1d_real_fill() {
    compile_ok(r#"
program test
    real :: a(4) = [1.0, 2.0, 3.0, 4.0]
    real :: b(4)
    b = eoshift(a, 1, boundary=0.0)
    print *, b(4)
end program test
"#);
}

#[test] fn eoshift_1d_char() {
    compile_ok(r#"
program test
    character(len=1) :: a(4) = ['a', 'b', 'c', 'd']
    character(len=1) :: b(4)
    b = eoshift(a, 1, boundary=' ')
    print *, b(4)
end program test
"#);
}

#[test] fn eoshift_1d_logical() {
    compile_ok(r#"
program test
    logical :: a(4) = [.true., .true., .false., .true.]
    logical :: b(4)
    b = eoshift(a, -1, boundary=.false.)
    print *, b(1)
end program test
"#);
}

#[test] fn eoshift_zero_shift() {
    compile_ok(r#"
program test
    integer :: a(4) = [1, 2, 3, 4]
    integer :: b(4)
    b = eoshift(a, 0)
    print *, b(1)
    print *, b(4)
end program test
"#);
}

// ── EOSHIFT — 2D with DIM ─────────────────────────────────────

#[test] fn eoshift_2d_dim1() {
    compile_ok(r#"
program test
    integer :: m(3,4) = reshape([(i, i=1,12)],[3,4])
    integer :: n(3,4)
    n = eoshift(m, 1, dim=1)
    print *, n(3,1)
end program test
"#);
}

#[test] fn eoshift_2d_dim2() {
    compile_ok(r#"
program test
    integer :: m(3,4) = reshape([(i, i=1,12)],[3,4])
    integer :: n(3,4)
    n = eoshift(m, 1, dim=2)
    print *, n(1,4)
end program test
"#);
}

#[test] fn eoshift_2d_with_boundary() {
    compile_ok(r#"
program test
    integer :: m(2,4) = reshape([1,2,3,4,5,6,7,8],[2,4])
    integer :: n(2,4)
    n = eoshift(m, 2, boundary=-99, dim=2)
    print *, n(1,3)
    print *, n(1,4)
end program test
"#);
}

#[test] fn eoshift_2d_row_shift_array() {
    compile_ok(r#"
program test
    integer :: m(3,3) = reshape([1,2,3,4,5,6,7,8,9],[3,3])
    integer :: shifts(3) = [0, 1, 2]
    integer :: n(3,3)
    n = eoshift(m, shifts, dim=2)
    print *, n(1,1)
end program test
"#);
}

// ── Combining CSHIFT and EOSHIFT with other operations ────────

#[test] fn cshift_in_expression() {
    compile_ok(r#"
program test
    integer :: a(4) = [1, 2, 3, 4]
    integer :: b(4)
    b = a + cshift(a, 1)
    print *, b(1)
end program test
"#);
}

#[test] fn eoshift_sum_pattern() {
    compile_ok(r#"
program test
    real :: a(5) = [1.0, 2.0, 3.0, 4.0, 5.0]
    real :: forward(5), backward(5), centered(5)
    forward  = eoshift(a,  1, boundary=0.0)
    backward = eoshift(a, -1, boundary=0.0)
    centered = (forward + backward) * 0.5
    print *, centered(3)
end program test
"#);
}
