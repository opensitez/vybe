! vybe-test: fortran/types/derived_type_extends_runtime
! origin: languages/fortran/tests/fortran/test_types.rs

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
    if ((nint(c%area + c%radius)) /= 15) then
    print *, "FAIL: want [15] got [", nint(c%area + c%radius), "]"
    stop 1
end if
end program test
