! vybe-test: fortran/arithmetic/compound_expression
! origin: languages/fortran/tests/fortran/test_arithmetic.rs
program t
if ((2 + 3 * 4) /= 14) then
    print *, "FAIL: want [14] got [", 2 + 3 * 4, "]"
    stop 1
end if
end program t
