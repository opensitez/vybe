! vybe-test: fortran/types/derived_type_extends
! origin: languages/fortran/tests/fortran/test_types.rs

program test
    type :: Shape
        real :: area
    end type Shape
    type, extends(Shape) :: Circle
        real :: radius
    end type Circle
    type(Circle) :: c
    c%radius = 5.0
    if ((c%radius) /= 5) then
    print *, "FAIL: want [5] got [", c%radius, "]"
    stop 1
end if
end program test
