! vybe-test: fortran/arithmetic/parenthesized
! origin: languages/fortran/tests/fortran/test_arithmetic.rs
program t
if (((2 + 3) * 4) /= 20) then
    print *, "FAIL: want [20] got [", (2 + 3) * 4, "]"
    stop 1
end if
end program t
