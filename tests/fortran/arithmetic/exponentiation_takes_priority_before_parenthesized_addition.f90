! vybe-test: fortran/arithmetic/exponentiation_takes_priority_before_parenthesized_addition
! origin: languages/fortran/tests/fortran/test_arithmetic.rs
program t
if (((2 + 1) ** 3) /= 27) then
    print *, "FAIL: want [27] got [", (2 + 1) ** 3, "]"
    stop 1
end if
end program t
