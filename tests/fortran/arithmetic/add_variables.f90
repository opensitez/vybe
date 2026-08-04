! vybe-test: fortran/arithmetic/add_variables
! origin: languages/fortran/tests/fortran/test_arithmetic.rs
program t
integer :: a = 10, b = 20
if ((a + b) /= 30) then
    print *, "FAIL: want [30] got [", a + b, "]"
    stop 1
end if
end program t
