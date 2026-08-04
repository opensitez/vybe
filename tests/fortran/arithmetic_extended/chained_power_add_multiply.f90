! vybe-test: fortran/arithmetic_extended/chained_power_add_multiply
! origin: languages/fortran/tests/fortran/test_arithmetic_extended.rs
program t
if ((3 ** 2 + 2 * 5) /= 19) then
    print *, "FAIL: want [19] got [", 3 ** 2 + 2 * 5, "]"
    stop 1
end if
end program t
