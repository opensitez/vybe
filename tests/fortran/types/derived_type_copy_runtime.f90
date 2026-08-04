! vybe-test: fortran/types/derived_type_copy_runtime
! origin: languages/fortran/tests/fortran/test_types.rs

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
    if ((nint(b%x + b%y)) /= 3) then
    print *, "FAIL: want [3] got [", nint(b%x + b%y), "]"
    stop 1
end if
end program test
