! vybe-test: fortran/types/derived_type_basic
! origin: languages/fortran/tests/fortran/test_types.rs

program test
    type :: Point
        real :: x
        real :: y
    end type Point
    type(Point) :: p
    p%x = 3.0
    p%y = 4.0
    if ((p%x + p%y) /= 7) then
    print *, "FAIL: want [7] got [", p%x + p%y, "]"
    stop 1
end if
end program test
