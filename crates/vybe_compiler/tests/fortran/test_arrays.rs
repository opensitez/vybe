use super::helpers::{compile_ok, run_prints};

// ── Fixed-size array declarations ─────────────────────────────

#[test]
fn array_dim_attr() {
    compile_ok("program t\n  integer, dimension(10) :: a\n  a(1) = 99\n  print *, a(1)\nend program t\n");
}

#[test]
fn array_shorthand() {
    compile_ok("program t\n  real :: v(3)\n  v(1) = 1.0\n  v(2) = 2.0\n  v(3) = 3.0\n  print *, v(2)\nend program t\n");
}

#[test]
fn array_2d() {
    compile_ok(r#"
program test
    integer :: m(3,3)
    integer :: i, j
    do i = 1, 3
        do j = 1, 3
            m(i,j) = i * 3 + j
        end do
    end do
    print *, m(2,2)
end program test
"#);
}

#[test]
fn array_3d() {
    compile_ok(r#"
program test
    integer :: t(2,2,2)
    t(1,1,1) = 111
    print *, t(1,1,1)
end program test
"#);
}

#[test]
fn array_init_loop() {
    let out = run_prints(r#"
program test
    integer :: a(5)
    integer :: i
    do i = 1, 5
        a(i) = i * i
    end do
    print *, a(3)
end program test
"#);
    assert_eq!(out, ["9"]);
}

#[test]
fn array_element_write() {
    let out = run_prints(r#"
program test
    integer :: a(3)
    a(1) = 10
    a(2) = 20
    a(3) = 30
    print *, a(1)
    print *, a(2)
    print *, a(3)
end program test
"#);
    assert_eq!(out, ["10", "20", "30"]);
}

// ── Allocatable arrays ────────────────────────────────────────

#[test]
fn alloc_1d_int() {
    compile_ok(r#"
program test
    integer, allocatable :: v(:)
    allocate(v(5))
    v(1) = 42
    print *, v(1)
    deallocate(v)
end program test
"#);
}

#[test]
fn alloc_1d_real() {
    compile_ok(r#"
program test
    real, allocatable :: v(:)
    allocate(v(3))
    v(1) = 3.14
    print *, v(1)
    deallocate(v)
end program test
"#);
}

#[test]
fn alloc_2d() {
    compile_ok(r#"
program test
    real, allocatable :: m(:,:)
    allocate(m(3,3))
    m(1,1) = 1.0
    print *, m(1,1)
    deallocate(m)
end program test
"#);
}

// ── Array slicing ─────────────────────────────────────────────

#[test]
fn slice_range() {
    compile_ok(r#"
program test
    integer :: a(5) = [1, 2, 3, 4, 5]
    integer :: b(3)
    b = a(2:4)
    print *, b(1)
end program test
"#);
}

#[test]
fn slice_step() {
    compile_ok(r#"
program test
    integer :: a(6) = [10, 20, 30, 40, 50, 60]
    integer :: b(3)
    b = a(1:6:2)
    print *, b(1)
end program test
"#);
}

#[test]
fn slice_from_start() {
    compile_ok(r#"
program test
    integer :: a(5) = [1, 2, 3, 4, 5]
    integer :: b(3)
    b = a(:3)
    print *, b(1)
end program test
"#);
}

#[test]
fn slice_to_end() {
    compile_ok(r#"
program test
    integer :: a(5) = [1, 2, 3, 4, 5]
    integer :: b(3)
    b = a(3:)
    print *, b(1)
end program test
"#);
}

// ── Array constructors ────────────────────────────────────────

#[test]
fn array_constructor_literal() {
    compile_ok("program t\n  integer :: a(3) = [1, 2, 3]\n  print *, a(2)\nend program t\n");
}

#[test]
fn array_constructor_old_syntax() {
    compile_ok("program t\n  integer :: a(3) = (/1, 2, 3/)\n  print *, a(1)\nend program t\n");
}

#[test]
fn array_constructor_implied_do() {
    compile_ok("program t\n  integer :: a(5) = [(i, i=1,5)]\n  print *, a(3)\nend program t\n");
}

// ── Array intrinsics ──────────────────────────────────────────

#[test]
fn intrinsic_sum() {
    compile_ok(r#"
program test
    integer :: a(5) = [1, 2, 3, 4, 5]
    print *, sum(a)
end program test
"#);
}

#[test]
fn intrinsic_product() {
    compile_ok(r#"
program test
    integer :: a(4) = [1, 2, 3, 4]
    print *, product(a)
end program test
"#);
}

#[test]
fn intrinsic_maxval() {
    compile_ok(r#"
program test
    integer :: a(5) = [3, 1, 4, 1, 5]
    print *, maxval(a)
end program test
"#);
}

#[test]
fn intrinsic_minval() {
    compile_ok(r#"
program test
    integer :: a(5) = [3, 1, 4, 1, 5]
    print *, minval(a)
end program test
"#);
}

#[test]
fn intrinsic_maxloc() {
    compile_ok(r#"
program test
    integer :: a(5) = [3, 1, 9, 1, 5]
    integer :: loc(1)
    loc = maxloc(a)
    print *, loc(1)
end program test
"#);
}

#[test]
fn intrinsic_minloc() {
    compile_ok(r#"
program test
    integer :: a(5) = [3, 1, 9, 1, 5]
    integer :: loc(1)
    loc = minloc(a)
    print *, loc(1)
end program test
"#);
}

#[test]
fn intrinsic_size() {
    compile_ok(r#"
program test
    integer :: a(7) = [1,2,3,4,5,6,7]
    print *, size(a)
end program test
"#);
}

#[test]
fn intrinsic_shape() {
    compile_ok(r#"
program test
    integer :: a(4)
    integer :: s(1)
    s = shape(a)
    print *, s(1)
end program test
"#);
}

#[test]
fn intrinsic_lbound() {
    compile_ok("program t\n  integer :: a(3)\n  print *, lbound(a, 1)\nend program t\n");
}

#[test]
fn intrinsic_ubound() {
    compile_ok("program t\n  integer :: a(3)\n  print *, ubound(a, 1)\nend program t\n");
}

#[test]
fn intrinsic_count() {
    compile_ok(r#"
program test
    logical :: mask(5) = [.true., .false., .true., .true., .false.]
    print *, count(mask)
end program test
"#);
}

#[test]
fn intrinsic_any() {
    compile_ok(r#"
program test
    logical :: mask(3) = [.false., .true., .false.]
    print *, any(mask)
end program test
"#);
}

#[test]
fn intrinsic_all() {
    compile_ok(r#"
program test
    logical :: mask(3) = [.true., .true., .true.]
    print *, all(mask)
end program test
"#);
}

// ── WHERE construct ───────────────────────────────────────────

#[test]
fn where_basic() {
    compile_ok(r#"
program test
    integer :: a(5) = [1, -2, 3, -4, 5]
    integer :: b(5)
    where (a > 0)
        b = a
    elsewhere
        b = 0
    end where
    print *, b(1)
end program test
"#);
}

#[test]
fn where_no_elsewhere() {
    compile_ok(r#"
program test
    real :: a(4) = [1.0, -2.0, 3.0, -4.0]
    where (a < 0.0)
        a = 0.0
    end where
    print *, a(1)
end program test
"#);
}

#[test]
fn where_set_to_zero() {
    compile_ok(r#"
program test
    integer :: v(6) = [1, -2, 3, -4, 5, -6]
    where (v < 0)
        v = 0
    end where
    print *, v(1)
end program test
"#);
}

// ── FORALL construct ──────────────────────────────────────────

#[test]
fn forall_basic() {
    compile_ok(r#"
program test
    integer :: a(5)
    forall (i = 1:5)
        a(i) = i * i
    end forall
    print *, a(3)
end program test
"#);
}

#[test]
fn forall_2d() {
    compile_ok(r#"
program test
    real :: m(3,3)
    forall (i = 1:3, j = 1:3)
        m(i,j) = real(i) + real(j)
    end forall
    print *, m(1,1)
end program test
"#);
}

// ── RESHAPE ──────────────────────────────────────────────────

#[test]
fn reshape_1d_to_2d() {
    compile_ok(r#"
program test
    integer :: a(6) = [1, 2, 3, 4, 5, 6]
    integer :: m(2,3)
    m = reshape(a, [2, 3])
    print *, m(1,1)
end program test
"#);
}

// ── MATMUL / DOT_PRODUCT ──────────────────────────────────────

#[test]
fn dot_product_basic() {
    compile_ok(r#"
program test
    integer :: a(3) = [1, 2, 3]
    integer :: b(3) = [4, 5, 6]
    print *, dot_product(a, b)
end program test
"#);
}

#[test]
fn matmul_basic() {
    compile_ok(r#"
program test
    integer :: a(2,2) = reshape([1,2,3,4],[2,2])
    integer :: b(2,2) = reshape([5,6,7,8],[2,2])
    integer :: c(2,2)
    c = matmul(a, b)
    print *, c(1,1)
end program test
"#);
}

// ── TRANSPOSE ────────────────────────────────────────────────

#[test]
fn transpose_basic() {
    compile_ok(r#"
program test
    integer :: a(2,3) = reshape([1,2,3,4,5,6],[2,3])
    integer :: b(3,2)
    b = transpose(a)
    print *, b(1,1)
end program test
"#);
}

// ── Passing arrays to subroutines ────────────────────────────

#[test]
fn array_subroutine_arg() {
    compile_ok(r#"
program test
    integer :: a(5) = [1, 2, 3, 4, 5]
    call print_first(a)
contains
    subroutine print_first(v)
        integer, intent(in) :: v(:)
        print *, v(1)
    end subroutine
end program test
"#);
}

#[test]
fn array_function_result() {
    compile_ok(r#"
program test
    integer :: a(3) = [10, 20, 30]
    print *, total(a)
contains
    function total(v) result(s)
        integer, intent(in) :: v(:)
        integer :: s
        s = sum(v)
    end function
end program test
"#);
}
