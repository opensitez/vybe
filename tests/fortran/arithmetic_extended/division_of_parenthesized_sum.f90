! vybe-test: fortran/arithmetic_extended/division_of_parenthesized_sum
! origin: languages/fortran/tests/fortran/test_arithmetic_extended.rs
program t
if ((48 / (6 + 2)) /= 6) then
    print *, "FAIL: want [6] got [", 48 / (6 + 2), "]"
    stop 1
end if
end program t
