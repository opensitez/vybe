use super::helpers::compile_ok;

// ── FORALL with mask (scalar condition) ───────────────────────

#[test] fn forall_mask_positive() {
    compile_ok(r#"
program test
    integer :: a(10) = [(i - 5, i=1,10)]
    integer :: b(10)
    b = 0
    forall (i = 1:10, a(i) > 0)
        b(i) = a(i)
    end forall
    print *, b(8)
    print *, b(3)
end program test
"#);
}

#[test] fn forall_mask_even() {
    compile_ok(r#"
program test
    integer :: a(10)
    a = 0
    forall (i = 1:10, mod(i, 2) == 0)
        a(i) = i
    end forall
    print *, a(4)
    print *, a(3)
end program test
"#);
}

#[test] fn forall_mask_odd() {
    compile_ok(r#"
program test
    integer :: a(6) = 0
    forall (i = 1:6, mod(i, 2) /= 0)
        a(i) = i * i
    end forall
    print *, a(1)
    print *, a(3)
    print *, a(2)
end program test
"#);
}

#[test] fn forall_mask_diagonal() {
    compile_ok(r#"
program test
    integer :: m(5,5)
    m = 0
    forall (i = 1:5, j = 1:5, i == j)
        m(i,j) = 1
    end forall
    print *, m(1,1)
    print *, m(2,2)
    print *, m(1,2)
end program test
"#);
}

#[test] fn forall_mask_upper_triangle() {
    compile_ok(r#"
program test
    integer :: m(4,4)
    m = 0
    forall (i = 1:4, j = 1:4, j > i)
        m(i,j) = i * 10 + j
    end forall
    print *, m(1,2)
    print *, m(1,1)
    print *, m(2,4)
end program test
"#);
}

#[test] fn forall_mask_lower_triangle() {
    compile_ok(r#"
program test
    real :: m(4,4)
    m = 0.0
    forall (i = 1:4, j = 1:4, j < i)
        m(i,j) = real(i + j)
    end forall
    print *, m(3,1)
    print *, m(1,2)
end program test
"#);
}

// ── FORALL with stride ────────────────────────────────────────

#[test] fn forall_stride_2() {
    compile_ok(r#"
program test
    integer :: a(10) = 0
    forall (i = 1:10:2)
        a(i) = i
    end forall
    print *, a(1)
    print *, a(3)
    print *, a(2)
end program test
"#);
}

#[test] fn forall_stride_3() {
    compile_ok(r#"
program test
    integer :: a(12) = 0
    forall (i = 1:12:3)
        a(i) = i
    end forall
    print *, a(1)
    print *, a(4)
    print *, a(2)
end program test
"#);
}

#[test] fn forall_2d_stride() {
    compile_ok(r#"
program test
    integer :: m(6,6) = 0
    forall (i = 1:6:2, j = 1:6:2)
        m(i,j) = i * j
    end forall
    print *, m(1,1)
    print *, m(3,3)
    print *, m(2,2)
end program test
"#);
}

// ── FORALL with mask and stride combined ──────────────────────

#[test] fn forall_stride_with_mask() {
    compile_ok(r#"
program test
    integer :: a(20) = 0
    forall (i = 1:20:2, i > 10)
        a(i) = i
    end forall
    print *, a(11)
    print *, a(9)
end program test
"#);
}

// ── FORALL with multiple assignments ──────────────────────────

#[test] fn forall_multiple_assignments() {
    compile_ok(r#"
program test
    real :: x(5), y(5), z(5)
    x = [1.0, 2.0, 3.0, 4.0, 5.0]
    y = [5.0, 4.0, 3.0, 2.0, 1.0]
    forall (i = 1:5)
        z(i) = x(i) + y(i)
        x(i) = x(i) * 2.0
    end forall
    print *, z(1)
    print *, x(1)
end program test
"#);
}

#[test] fn forall_symmetrize_matrix() {
    compile_ok(r#"
program test
    real :: m(3,3)
    m = 0.0
    m(1,2) = 5.0; m(1,3) = 7.0; m(2,3) = 9.0
    forall (i = 1:3, j = 1:3, i /= j)
        m(i,j) = m(i,j) + m(j,i)
    end forall
    print *, m(2,1)
end program test
"#);
}

// ── FORALL with elemental function in body ────────────────────

#[test] fn forall_elemental_call() {
    compile_ok(r#"
program test
    real :: a(5) = [1.0, 4.0, 9.0, 16.0, 25.0]
    real :: b(5)
    forall (i = 1:5)
        b(i) = sqrt(a(i))
    end forall
    print *, b(1)
    print *, b(4)
end program test
"#);
}

#[test] fn forall_with_abs() {
    compile_ok(r#"
program test
    integer :: a(6) = [-3, 2, -1, 4, -5, 0]
    integer :: b(6)
    forall (i = 1:6)
        b(i) = abs(a(i))
    end forall
    print *, b(1)
    print *, b(2)
end program test
"#);
}

// ── Nested FORALL ─────────────────────────────────────────────

#[test] fn nested_forall() {
    compile_ok(r#"
program test
    integer :: m(3,3)
    m = 0
    forall (i = 1:3)
        forall (j = 1:3)
            m(i,j) = i * 10 + j
        end forall
    end forall
    print *, m(2,3)
end program test
"#);
}

#[test] fn nested_forall_with_mask() {
    compile_ok(r#"
program test
    integer :: m(4,4) = 0
    forall (i = 1:4, i <= 2)
        forall (j = 1:4, j > i)
            m(i,j) = i + j
        end forall
    end forall
    print *, m(1,2)
    print *, m(1,1)
end program test
"#);
}

// ── FORALL initialization patterns ────────────────────────────

#[test] fn forall_identity_matrix() {
    compile_ok(r#"
program test
    real :: id(5,5)
    id = 0.0
    forall (i = 1:5)
        id(i,i) = 1.0
    end forall
    print *, id(1,1)
    print *, id(1,2)
end program test
"#);
}

#[test] fn forall_tridiagonal() {
    compile_ok(r#"
program test
    real :: m(5,5)
    m = 0.0
    forall (i = 1:5)
        m(i,i) = 2.0
    end forall
    forall (i = 1:4)
        m(i,i+1) = -1.0
        m(i+1,i) = -1.0
    end forall
    print *, m(1,1)
    print *, m(1,2)
end program test
"#);
}

#[test] fn forall_outer_product() {
    compile_ok(r#"
program test
    real :: u(4) = [1.0, 2.0, 3.0, 4.0]
    real :: v(4) = [1.0, 2.0, 3.0, 4.0]
    real :: m(4,4)
    forall (i = 1:4, j = 1:4)
        m(i,j) = u(i) * v(j)
    end forall
    print *, m(2,3)
end program test
"#);
}

// ── Single-line FORALL (statement form) ───────────────────────

#[test] fn forall_statement_form() {
    compile_ok(r#"
program test
    real :: a(5)
    forall (i = 1:5) a(i) = real(i) ** 2
    print *, a(3)
end program test
"#);
}

#[test] fn forall_statement_with_mask() {
    compile_ok(r#"
program test
    integer :: a(10) = 0
    forall (i = 1:10, mod(i,3) == 0) a(i) = i
    print *, a(3)
    print *, a(6)
    print *, a(4)
end program test
"#);
}

// ── FORALL in subroutines / modules ───────────────────────────

#[test] fn forall_in_subroutine() {
    compile_ok(r#"
program test
    real :: a(5)
    call fill_squares(a)
    print *, a(3)
contains
    subroutine fill_squares(x)
        real, intent(out) :: x(:)
        integer :: n
        n = size(x)
        forall (i = 1:n)
            x(i) = real(i) ** 2
        end forall
    end subroutine fill_squares
end program test
"#);
}
