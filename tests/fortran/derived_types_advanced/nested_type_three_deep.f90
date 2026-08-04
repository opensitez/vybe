! vybe-test: fortran/derived_types_advanced/nested_type_three_deep
! origin: languages/fortran/tests/fortran/test_derived_types_advanced.rs

program test
    type :: Coord
        real :: val
    end type Coord
    type :: Point
        type(Coord) :: x, y
    end type Point
    type :: Segment
        type(Point) :: start, finish
    end type Segment
    type(Segment) :: s
    s%start%x%val = 1.0
    print *, s%start%x%val
end program test
