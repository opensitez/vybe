! vybe-test: fortran/associate_construct_extended/associate_dtype_real_field
! origin: languages/fortran/tests/fortran/test_associate_construct_extended.rs
program t
type :: Point
real :: x, y
end type Point
type(Point) :: pt
pt%x = 6.0
pt%y = 8.0
associate (abscissa => pt%x)
if ((int(abscissa)) /= 6) then
    print *, "FAIL: want [6] got [", int(abscissa), "]"
    stop 1
end if
end associate
end program t
