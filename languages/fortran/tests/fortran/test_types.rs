use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// Fortran: Type declarations, derived types, modules
// ═══════════════════════════════════════════════════════════

#[test]
fn derived_type_basic() {
    let out = run_prints(
        r#"
program test
    type :: Point
        real :: x
        real :: y
    end type Point
    type(Point) :: p
    p%x = 3.0
    p%y = 4.0
    print *, p%x + p%y
end program test
"#,
    );
    assert_eq!(out, vec!["7"]);
}

#[test]
fn derived_type_basic_runtime() {
    let out = run_prints(
        r#"
program test
    type :: Point
        real :: x
        real :: y
    end type Point
    type(Point) :: p
    p%x = 3.0
    p%y = 4.0
    print *, nint(p%x + p%y)
end program test
"#,
    );
    assert_eq!(out, vec!["7"]);
}

#[test]
fn derived_type_with_methods() {
    let out = run_prints(
        r#"
program test
    type :: Counter
        integer :: value = 0
    contains
        procedure :: increment
    end type Counter
    type(Counter) :: c
    print *, c%value
contains
    subroutine increment(self)
        class(Counter), intent(inout) :: self
        self%value = self%value + 1
    end subroutine increment
end program test
"#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn derived_type_with_methods_runtime() {
    let out = run_prints(
        r#"
program test
    type :: Counter
        integer :: value = 0
    contains
        procedure :: increment
    end type Counter
    type(Counter) :: c
    call c%increment()
    call c%increment()
    print *, c%value
contains
    subroutine increment(self)
        class(Counter), intent(inout) :: self
        self%value = self%value + 1
    end subroutine increment
end program test
"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn module_basic() {
    let out = run_prints(
        r#"
module constants
    real, parameter :: PI = 3.14159
    real, parameter :: E = 2.71828
end module constants

program test
    use constants
    print *, PI
end program test
"#,
    );
    assert_eq!(out[0], "3.14159");
}

#[test]
fn module_with_function() {
    let out = run_prints(
        r#"
module math_utils
    implicit none
contains
    function square(x) result(res)
        real, intent(in) :: x
        real :: res
        res = x * x
    end function square
end module math_utils

program test
    use math_utils
    print *, square(5.0)
end program test
"#,
    );
    assert_eq!(out, vec!["25"]);
}

#[test]
fn module_with_function_runtime() {
    let out = run_prints(
        r#"
module math_utils
    implicit none
contains
    function square(x) result(res)
        real, intent(in) :: x
        real :: res
        res = x * x
    end function square
end module math_utils

program test
    use math_utils
    print *, nint(square(5.0))
end program test
"#,
    );
    assert_eq!(out, vec!["25"]);
}

#[test]
fn integer_types() {
    let out = run_prints(
        r#"
program test
    integer :: a = 10
    real :: b = 3.14
    double precision :: c = 2.718281828
    logical :: d = .true.
    character(len=10) :: e = "hello"
    print *, a
    print *, b
    print *, d
    print *, e
end program test
"#,
    );
    assert_eq!(out[0], "10");
    assert_eq!(out[1], "3.14");
}

#[test]
fn array_declaration() {
    let out = run_prints(
        r#"
program test
    integer, dimension(5) :: arr
    integer :: i
    do i = 1, 5
        arr(i) = i * 10
    end do
    print *, arr(3)
end program test
"#,
    );
    assert_eq!(out, vec!["30"]);
}

#[test]
fn array_declaration_runtime_indexed_assignment() {
    let out = run_prints(
        r#"
program test
    integer, dimension(5) :: arr
    integer :: i
    do i = 1, 5
        arr(i) = i * 10
    end do
    print *, arr(1)
    print *, arr(5)
end program test
"#,
    );
    assert_eq!(out, vec!["10", "50"]);
}

#[test]
fn allocatable_array() {
    let out = run_prints(
        r#"
program test
    integer, allocatable :: arr(:)
    allocate(arr(10))
    arr(1) = 42
    print *, arr(1)
    deallocate(arr)
end program test
"#,
    );
    assert_eq!(out, vec!["42"]);
}

#[test]
fn derived_type_extends() {
    let out = run_prints(
        r#"
program test
    type :: Shape
        real :: area
    end type Shape
    type, extends(Shape) :: Circle
        real :: radius
    end type Circle
    type(Circle) :: c
    c%radius = 5.0
    print *, c%radius
end program test
"#,
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn derived_type_extends_runtime() {
    let out = run_prints(
        r#"
program test
    type :: Shape
        real :: area
    end type Shape
    type, extends(Shape) :: Circle
        real :: radius
    end type Circle
    type(Circle) :: c
    c%area = 10.0
    c%radius = 5.0
    print *, nint(c%area + c%radius)
end program test
"#,
    );
    assert_eq!(out, vec!["15"]);
}

#[test]
fn derived_type_copy_runtime() {
    let out = run_prints(
        r#"
program test
    type :: Point
        real :: x
        real :: y
    end type Point
    type(Point) :: a
    type(Point) :: b
    a%x = 1.0
    a%y = 2.0
    b = a
    print *, nint(b%x + b%y)
end program test
"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn use_only_runtime() {
    let out = run_prints(
        r#"
module mymod
    integer :: x = 10
    integer :: y = 20
end module mymod

program test
    use mymod, only: x
    print *, x
end program test
"#,
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn use_only() {
    let out = run_prints(
        r#"
module mymod
    integer :: x = 10
    integer :: y = 20
end module mymod

program test
    use mymod, only: x
    print *, x
end program test
"#,
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn interface_basic() {
    let out = run_prints(
        r#"
program test
    interface
        function add(a, b) result(res)
            integer, intent(in) :: a, b
            integer :: res
        end function add
    end interface
    print *, "ok"
end program test
"#,
    );
    assert_eq!(out, vec!["ok"]);
}

#[test]
fn derived_type_procedure_component_decl() {
    let out = run_prints(
        r#"
program test
    implicit none
    abstract interface
        function rhs_func(t) result(v)
            real, intent(in) :: t
            real :: v
        end function rhs_func
    end interface

    type :: CallbackBox
        procedure(rhs_func), pointer, nopass :: fn
    end type CallbackBox

    print *, "ok"
end program test
"#,
    );
    assert_eq!(out, vec!["ok"]);
}
