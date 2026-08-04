! vybe-test: fortran/arithmetic/complex_arithmetic
! origin: languages/fortran/tests/fortran/test_arithmetic.rs
program t
if ((2 ** 3 + 1) /= 9) then
    print *, "FAIL: want [9] got [", 2 ** 3 + 1, "]"
    stop 1
end if
end program t
