use super::helpers::{compile_ok, run_prints};

// ── PURE functions ────────────────────────────────────────────

#[test]
fn pure_function_basic() {
    compile_ok(r#"
program test
    print *, square(5)
contains
    pure function square(x) result(res)
        integer, intent(in) :: x
        integer :: res
        res = x * x
    end function square
end program test
"#);
}

#[test]
fn pure_function_real() {
    compile_ok(r#"
program test
    print *, hyp(3.0, 4.0)
contains
    pure function hyp(a, b) result(c)
        real, intent(in) :: a, b
        real :: c
        c = sqrt(a*a + b*b)
    end function hyp
end program test
"#);
}

#[test]
fn pure_subroutine() {
    compile_ok(r#"
program test
    integer :: x = 5
    call double_it(x)
    print *, x
contains
    pure subroutine double_it(n)
        integer, intent(inout) :: n
        n = n * 2
    end subroutine double_it
end program test
"#);
}

#[test]
fn pure_in_module() {
    compile_ok(r#"
module math_pure
    implicit none
contains
    pure function add(a, b) result(res)
        integer, intent(in) :: a, b
        integer :: res
        res = a + b
    end function add
end module math_pure

program test
    use math_pure
    print *, add(3, 4)
end program test
"#);
}

// ── ELEMENTAL functions ───────────────────────────────────────

#[test]
fn elemental_function_basic() {
    compile_ok(r#"
program test
    integer :: a(3) = [1, 2, 3]
    integer :: b(3)
    b = double_elem(a)
    print *, b(1)
contains
    elemental function double_elem(x) result(res)
        integer, intent(in) :: x
        integer :: res
        res = x * 2
    end function double_elem
end program test
"#);
}

#[test]
fn elemental_on_scalar() {
    compile_ok(r#"
program test
    print *, cube(3)
contains
    elemental function cube(x) result(res)
        integer, intent(in) :: x
        integer :: res
        res = x * x * x
    end function cube
end program test
"#);
}

#[test]
fn elemental_real() {
    compile_ok(r#"
program test
    real :: a(4) = [1.0, 4.0, 9.0, 16.0]
    real :: b(4)
    b = root(a)
    print *, b(1)
contains
    elemental function root(x) result(r)
        real, intent(in) :: x
        real :: r
        r = sqrt(x)
    end function root
end program test
"#);
}

#[test]
fn elemental_subroutine() {
    compile_ok(r#"
program test
    integer :: a(3) = [1, 2, 3]
    call negate(a)
    print *, a(1)
contains
    elemental subroutine negate(x)
        integer, intent(inout) :: x
        x = -x
    end subroutine negate
end program test
"#);
}

// ── RECURSIVE functions (more cases) ─────────────────────────

#[test]
fn recursive_power() {
    compile_ok(r#"
program test
    print *, power(2, 8)
contains
    recursive function power(base, exp) result(res)
        integer, intent(in) :: base, exp
        integer :: res
        if (exp == 0) then
            res = 1
        else
            res = base * power(base, exp - 1)
        end if
    end function power
end program test
"#);
}

#[test]
fn recursive_gcd() {
    compile_ok(r#"
program test
    print *, gcd(48, 18)
contains
    recursive function gcd(a, b) result(res)
        integer, intent(in) :: a, b
        integer :: res
        if (b == 0) then
            res = a
        else
            res = gcd(b, mod(a, b))
        end if
    end function gcd
end program test
"#);
}

#[test]
fn recursive_ackermann() {
    compile_ok(r#"
program test
    print *, ack(2, 3)
contains
    recursive function ack(m, n) result(res)
        integer, intent(in) :: m, n
        integer :: res
        if (m == 0) then
            res = n + 1
        else if (n == 0) then
            res = ack(m - 1, 1)
        else
            res = ack(m - 1, ack(m, n - 1))
        end if
    end function ack
end program test
"#);
}

#[test]
fn recursive_sum() {
    compile_ok(r#"
program test
    print *, rsum(10)
contains
    recursive function rsum(n) result(res)
        integer, intent(in) :: n
        integer :: res
        if (n <= 0) then
            res = 0
        else
            res = n + rsum(n - 1)
        end if
    end function rsum
end program test
"#);
}

// ── OPTIONAL arguments ────────────────────────────────────────

#[test]
fn optional_basic() {
    compile_ok(r#"
program test
    call greet('Alice')
    call greet('Bob', 'Dr.')
contains
    subroutine greet(name, title)
        character(len=*), intent(in) :: name
        character(len=*), intent(in), optional :: title
        if (present(title)) then
            print *, trim(title) // ' ' // trim(name)
        else
            print *, trim(name)
        end if
    end subroutine greet
end program test
"#);
}

#[test]
fn optional_integer() {
    compile_ok(r#"
program test
    print *, with_default(5)
    print *, with_default(5, 10)
contains
    function with_default(x, y) result(res)
        integer, intent(in) :: x
        integer, intent(in), optional :: y
        integer :: res
        if (present(y)) then
            res = x + y
        else
            res = x
        end if
    end function with_default
end program test
"#);
}

#[test]
fn optional_present_check() {
    compile_ok(r#"
program test
    call maybe(3)
contains
    subroutine maybe(n, extra)
        integer, intent(in) :: n
        integer, intent(in), optional :: extra
        integer :: total
        total = n
        if (present(extra)) total = total + extra
        print *, total
    end subroutine maybe
end program test
"#);
}

#[test]
fn optional_present_check_runtime() {
    let out = run_prints(r#"
program test
    call maybe(3)
    call maybe(3, 4)
contains
    subroutine maybe(n, extra)
        integer, intent(in) :: n
        integer, intent(in), optional :: extra
        integer :: total
        total = n
        if (present(extra)) total = total + extra
        print *, total
    end subroutine maybe
end program test
"#);
    assert_eq!(out, vec!["3", "7"]);
}

// ── INTENT (more cases) ───────────────────────────────────────

#[test]
fn intent_in_multiple() {
    compile_ok(r#"
program test
    print *, add3(1, 2, 3)
contains
    function add3(a, b, c) result(res)
        integer, intent(in) :: a, b, c
        integer :: res
        res = a + b + c
    end function add3
end program test
"#);
}

#[test]
fn intent_out_scalar() {
    compile_ok(r#"
program test
    integer :: x
    call set_value(x)
    print *, x
contains
    subroutine set_value(n)
        integer, intent(out) :: n
        n = 42
    end subroutine set_value
end program test
"#);
}

#[test]
fn intent_inout_double() {
    compile_ok(r#"
program test
    integer :: x = 5
    call double_it(x)
    print *, x
contains
    subroutine double_it(n)
        integer, intent(inout) :: n
        n = n * 2
    end subroutine double_it
end program test
"#);
}

// ── CONTAINS in program ───────────────────────────────────────

#[test]
fn contains_multiple_funcs() {
    compile_ok(r#"
program test
    print *, add(3, 4)
    print *, mul(3, 4)
contains
    function add(a, b) result(r)
        integer, intent(in) :: a, b
        integer :: r
        r = a + b
    end function add

    function mul(a, b) result(r)
        integer, intent(in) :: a, b
        integer :: r
        r = a * b
    end function mul
end program test
"#);
}

#[test]
fn contains_subroutine_calls_function() {
    compile_ok(r#"
program test
    call run()
contains
    function compute(x) result(r)
        integer, intent(in) :: x
        integer :: r
        r = x * x + 1
    end function compute

    subroutine run()
        print *, compute(4)
    end subroutine run
end program test
"#);
}
