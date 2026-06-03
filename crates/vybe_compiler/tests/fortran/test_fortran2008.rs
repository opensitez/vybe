use super::helpers::{compile_ok, run_prints};

// ── BLOCK construct ───────────────────────────────────────────

#[test]
fn block_basic() {
    compile_ok(
        r#"
program test
    integer :: x = 5
    block
        integer :: temp
        temp = x * 2
        print *, temp
    end block
end program test
"#,
    );
}

#[test]
fn block_local_scope() {
    compile_ok(
        r#"
program test
    integer :: i = 10
    block
        integer :: i   ! shadows outer i
        i = 99
        print *, i
    end block
    print *, i  ! outer i unchanged
end program test
"#,
    );
}

#[test]
fn block_swap() {
    compile_ok(
        r#"
program test
    integer :: a = 3, b = 7
    block
        integer :: tmp
        tmp = a
        a = b
        b = tmp
    end block
    print *, a
    print *, b
end program test
"#,
    );
}

#[test]
fn block_allocatable() {
    compile_ok(
        r#"
program test
    block
        integer, allocatable :: arr(:)
        allocate(arr(5))
        arr = [1, 2, 3, 4, 5]
        print *, sum(arr)
    end block
end program test
"#,
    );
}

#[test]
fn block_nested() {
    compile_ok(
        r#"
program test
    integer :: x = 1
    block
        integer :: y
        y = x + 10
        block
            integer :: z
            z = y + 100
            print *, z
        end block
    end block
end program test
"#,
    );
}

// ── DO CONCURRENT (Fortran 2008) ──────────────────────────────

#[test]
fn do_concurrent_basic() {
    let out = run_prints(
        r#"
program test
    integer :: a(10)
    do concurrent (i = 1:10)
        a(i) = i * i
    end do
    print *, a(3)
end program test
"#,
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn do_concurrent_2d() {
    let out = run_prints(
        r#"
program test
    real :: m(4,4)
    do concurrent (i = 1:4, j = 1:4)
        m(i,j) = real(i) * real(j)
    end do
    print *, m(2,3)
end program test
"#,
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn do_concurrent_mask() {
    let out = run_prints(
        r#"
program test
    integer :: a(10)
    a = 0
    do concurrent (i = 1:10, mod(i, 2) == 0)
        a(i) = i
    end do
    print *, a(4)
end program test
"#,
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn do_concurrent_locality() {
    compile_ok(
        r#"
program test
    integer :: a(5), b(5)
    b = [1, 2, 3, 4, 5]
    do concurrent (i = 1:5) local(tmp)
        integer :: tmp
        tmp = b(i) * 2
        a(i) = tmp
    end do
    print *, a(3)
end program test
"#,
    );
}

#[test]
fn do_concurrent_shared() {
    compile_ok(
        r#"
program test
    integer :: a(5)
    integer :: factor
    factor = 3
    do concurrent (i = 1:5) shared(factor)
        a(i) = i * factor
    end do
    print *, a(2)
end program test
"#,
    );
}

// ── CONTIGUOUS attribute ──────────────────────────────────────

#[test]
fn contiguous_pointer() {
    compile_ok(
        r#"
program test
    integer, target :: a(10) = [(i, i=1,10)]
    integer, pointer, contiguous :: p(:)
    p => a
    print *, p(3)
end program test
"#,
    );
}

#[test]
fn contiguous_dummy() {
    compile_ok(
        r#"
program test
    integer :: a(5) = [1, 2, 3, 4, 5]
    call process(a)
contains
    subroutine process(v)
        integer, intent(in), contiguous :: v(:)
        print *, v(1)
    end subroutine process
end program test
"#,
    );
}

#[test]
fn is_contiguous() {
    compile_ok(
        r#"
program test
    integer, target :: a(10)
    integer, pointer :: p(:)
    p => a
    print *, is_contiguous(p)
end program test
"#,
    );
}

// ── SUBMODULE (Fortran 2008) ──────────────────────────────────

#[test]
fn submodule_basic() {
    compile_ok(
        r#"
module parent_mod
    implicit none
    interface
        module function compute(x) result(r)
            integer, intent(in) :: x
            integer :: r
        end function compute
    end interface
end module parent_mod

submodule (parent_mod) parent_mod_impl
    implicit none
contains
    module function compute(x) result(r)
        integer, intent(in) :: x
        integer :: r
        r = x * x
    end function compute
end submodule parent_mod_impl

program test
    use parent_mod
    print *, compute(5)
end program test
"#,
    );
}

// ── IMPURE ELEMENTAL ──────────────────────────────────────────

#[test]
fn impure_elemental_basic() {
    compile_ok(
        r#"
program test
    integer :: a(3) = [1, 2, 3]
    call print_elem(a)
contains
    impure elemental subroutine print_elem(x)
        integer, intent(in) :: x
        print *, x
    end subroutine print_elem
end program test
"#,
    );
}

#[test]
fn impure_elemental_function() {
    compile_ok(
        r#"
program test
    integer :: a(3) = [1, 2, 3]
    integer :: b(3)
    b = logged_double(a)
    print *, b(1)
contains
    impure elemental function logged_double(x) result(r)
        integer, intent(in) :: x
        integer :: r
        r = x * 2
    end function logged_double
end program test
"#,
    );
}

// ── Internal subprogram as actual argument ────────────────────

#[test]
fn internal_proc_as_arg() {
    let out = run_prints(
        r#"
program test
    integer :: result
    result = apply(3, double_it)
    print *, result
contains
    function apply(x, fn) result(r)
        integer, intent(in) :: x
        interface
            function fn(n) result(v)
                integer, intent(in) :: n
                integer :: v
            end function fn
        end interface
        integer :: r
        r = fn(x)
    end function apply

    function double_it(n) result(v)
        integer, intent(in) :: n
        integer :: v
        v = n * 2
    end function double_it
end program test
"#,
    );
    assert_eq!(out, vec!["6"]);
}

// ── FINDLOC (Fortran 2008) ────────────────────────────────────

#[test]
fn findloc_basic() {
    compile_ok(
        r#"
program test
    integer :: a(6) = [3, 1, 4, 1, 5, 9]
    integer :: loc(1)
    loc = findloc(a, 4)
    print *, loc(1)
end program test
"#,
    );
}

#[test]
fn findloc_not_found() {
    compile_ok(
        r#"
program test
    integer :: a(5) = [1, 2, 3, 4, 5]
    integer :: loc(1)
    loc = findloc(a, 99)
    print *, loc(1)
end program test
"#,
    );
}

#[test]
fn findloc_back() {
    compile_ok(
        r#"
program test
    integer :: a(6) = [1, 2, 1, 2, 1, 2]
    integer :: loc(1)
    loc = findloc(a, 1, back=.true.)
    print *, loc(1)
end program test
"#,
    );
}

// ── Non-default accessibility ─────────────────────────────────

#[test]
fn module_default_private() {
    compile_ok(
        r#"
module strict_mod
    implicit none
    private
    public :: visible

    integer :: hidden = 0
    integer, public :: visible = 42
end module strict_mod

program test
    use strict_mod
    print *, visible
end program test
"#,
    );
}

// ── IEEE arithmetic (Fortran 2003/2008) ──────────────────────

#[test]
fn ieee_arithmetic_use() {
    compile_ok(
        r#"
program test
    use ieee_arithmetic
    real :: x
    x = ieee_value(x, ieee_positive_inf)
    print *, ieee_is_finite(x)
end program test
"#,
    );
}

#[test]
fn ieee_nan() {
    compile_ok(
        r#"
program test
    use ieee_arithmetic
    real :: x
    x = ieee_value(x, ieee_quiet_nan)
    print *, ieee_is_nan(x)
end program test
"#,
    );
}

#[test]
fn ieee_exceptions() {
    compile_ok(
        r#"
program test
    use ieee_exceptions
    type(ieee_flag_type) :: flag
    logical :: halting
    call ieee_get_halting_mode(ieee_divide_by_zero, halting)
    print *, halting
end program test
"#,
    );
}

// ── PACK and UNPACK ───────────────────────────────────────────

#[test]
fn pack_basic() {
    compile_ok(
        r#"
program test
    integer :: a(5) = [1, 2, 3, 4, 5]
    logical :: mask(5) = [.true., .false., .true., .false., .true.]
    integer :: b(3)
    b = pack(a, mask)
    print *, b(1)
end program test
"#,
    );
}

#[test]
fn unpack_basic() {
    compile_ok(
        r#"
program test
    integer :: a(3) = [10, 20, 30]
    logical :: mask(5) = [.true., .false., .true., .false., .true.]
    integer :: b(5)
    integer :: fill(5) = [0, 0, 0, 0, 0]
    b = unpack(a, mask, fill)
    print *, b(1)
end program test
"#,
    );
}

#[test]
fn spread_intrinsic() {
    compile_ok(
        r#"
program test
    integer :: a(3) = [1, 2, 3]
    integer :: m(3, 4)
    m = spread(a, 2, 4)
    print *, m(2, 1)
end program test
"#,
    );
}
