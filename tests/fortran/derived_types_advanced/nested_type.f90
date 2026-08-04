! vybe-test: fortran/derived_types_advanced/nested_type
! origin: languages/fortran/tests/fortran/test_derived_types_advanced.rs

program test
    type :: Point
        real :: x, y
    end type Point
    type :: Circle
        type(Point) :: center
        real :: radius
    end type Circle
    type(Circle) :: c
    c%center%x = 0.0
    c%center%y = 0.0
    c%radius = 5.0
    print *, c%radius
end program test
