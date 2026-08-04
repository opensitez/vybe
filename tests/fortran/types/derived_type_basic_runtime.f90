! vybe-test: fortran/types/derived_type_basic_runtime
! origin: languages/fortran/tests/fortran/test_types.rs

program test
    type :: Point
        real :: x
        real :: y
    end type Point
    type(Point) :: p
    p%x = 3.0
    p%y = 4.0
    if ((nint(p%x + p%y)) /= 7) then
    print *, "FAIL: want [7] got [", nint(p%x + p%y), "]"
    stop 1
end if
end program test
