! vybe-test: fortran/arithmetic/multiply_integers
! origin: languages/fortran/tests/fortran/test_arithmetic.rs
program t
if ((6 * 7) /= 42) then
    print *, "FAIL: want [42] got [", 6 * 7, "]"
    stop 1
end if
end program t
