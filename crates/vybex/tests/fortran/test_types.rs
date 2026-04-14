use super::helpers::compile_ok;

// ═══════════════════════════════════════════════════════════
// Fortran: Type declarations, derived types, modules
// ═══════════════════════════════════════════════════════════

#[test]
fn derived_type_basic() {
    compile_ok(r#"
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
"#);
}

#[test]
fn derived_type_with_methods() {
    compile_ok(r#"
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
"#);
}

#[test]
fn module_basic() {
    compile_ok(r#"
module constants
    real, parameter :: PI = 3.14159
    real, parameter :: E = 2.71828
end module constants

program test
    use constants
    print *, PI
end program test
"#);
}

#[test]
fn module_with_function() {
    compile_ok(r#"
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
"#);
}

#[test]
fn integer_types() {
    compile_ok(r#"
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
"#);
}

#[test]
fn array_declaration() {
    compile_ok(r#"
program test
    integer, dimension(5) :: arr
    integer :: i
    do i = 1, 5
        arr(i) = i * 10
    end do
    print *, arr(3)
end program test
"#);
}

#[test]
fn allocatable_array() {
    compile_ok(r#"
program test
    integer, allocatable :: arr(:)
    allocate(arr(10))
    arr(1) = 42
    print *, arr(1)
    deallocate(arr)
end program test
"#);
}

#[test]
fn derived_type_extends() {
    compile_ok(r#"
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
"#);
}

#[test]
fn use_only() {
    compile_ok(r#"
module mymod
    integer :: x = 10
    integer :: y = 20
end module mymod

program test
    use mymod, only: x
    print *, x
end program test
"#);
}

#[test]
fn interface_basic() {
    compile_ok(r#"
program test
    interface
        function add(a, b) result(res)
            integer, intent(in) :: a, b
            integer :: res
        end function add
    end interface
    print *, "ok"
end program test
"#);
}
