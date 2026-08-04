! vybe-test: fortran/arithmetic/subtract_variables
! origin: languages/fortran/tests/fortran/test_arithmetic.rs
program t
integer :: a = 10, b = 3
if ((a - b) /= 7) then
    print *, "FAIL: want [7] got [", a - b, "]"
    stop 1
end if
end program t
