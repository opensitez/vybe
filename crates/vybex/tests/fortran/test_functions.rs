use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// Fortran: Functions and subroutines
// ═══════════════════════════════════════════════════════════

#[test]
fn subroutine_basic() {
    compile_ok(r#"
program test
    call greet()
contains
    subroutine greet()
        print *, "Hello"
    end subroutine greet
end program test
"#);
}

#[test]
fn function_basic() {
    compile_ok(r#"
program test
    integer :: result
    result = square(5)
    print *, result
contains
    function square(x) result(res)
        integer, intent(in) :: x
        integer :: res
        res = x * x
    end function square
end program test
"#);
}

#[test]
fn subroutine_with_args() {
    compile_ok(r#"
program test
    integer :: a, b, c
    a = 3
    b = 4
    call add_nums(a, b, c)
    print *, c
contains
    subroutine add_nums(x, y, result)
        integer, intent(in) :: x, y
        integer, intent(out) :: result
        result = x + y
    end subroutine add_nums
end program test
"#);
}

#[test]
fn function_with_return_type() {
    compile_ok(r#"
program test
    print *, cube(3)
contains
    integer function cube(n)
        integer, intent(in) :: n
        cube = n * n * n
    end function cube
end program test
"#);
}

#[test]
fn recursive_function() {
    compile_ok(r#"
program test
    print *, factorial(5)
contains
    recursive function factorial(n) result(res)
        integer, intent(in) :: n
        integer :: res
        if (n <= 1) then
            res = 1
        else
            res = n * factorial(n - 1)
        end if
    end function factorial
end program test
"#);
}

#[test]
fn multiple_functions() {
    compile_ok(r#"
program test
    print *, add(3, 4)
    print *, multiply(3, 4)
contains
    function add(a, b) result(res)
        integer, intent(in) :: a, b
        integer :: res
        res = a + b
    end function add
    function multiply(a, b) result(res)
        integer, intent(in) :: a, b
        integer :: res
        res = a * b
    end function multiply
end program test
"#);
}
