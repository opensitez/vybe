! vybe-test: fortran/derived_types_advanced/type_constructor_positional
! origin: languages/fortran/tests/fortran/test_derived_types_advanced.rs

program test
    type :: Point
        real :: x, y
    end type Point
    type(Point) :: p
    p = Point(3.0, 4.0)
    if ((p%x) /= 3) then
    print *, "FAIL: want [3] got [", p%x, "]"
    stop 1
end if
end program test
