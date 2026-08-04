! vybe-test: fortran/arithmetic_extended/mixed_add_multiply_subtract
! origin: languages/fortran/tests/fortran/test_arithmetic_extended.rs
program t
if ((2 + 3 * 4 - 5) /= 9) then
    print *, "FAIL: want [9] got [", 2 + 3 * 4 - 5, "]"
    stop 1
end if
end program t
