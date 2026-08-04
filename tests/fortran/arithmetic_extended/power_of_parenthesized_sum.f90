! vybe-test: fortran/arithmetic_extended/power_of_parenthesized_sum
! origin: languages/fortran/tests/fortran/test_arithmetic_extended.rs
program t
if (((2 + 1) ** 3) /= 27) then
    print *, "FAIL: want [27] got [", (2 + 1) ** 3, "]"
    stop 1
end if
end program t
