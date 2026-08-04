! vybe-test: fortran/arithmetic/multiply_variables
! origin: languages/fortran/tests/fortran/test_arithmetic.rs
program t
integer :: a = 6, b = 7
if ((a * b) /= 42) then
    print *, "FAIL: want [42] got [", a * b, "]"
    stop 1
end if
end program t
