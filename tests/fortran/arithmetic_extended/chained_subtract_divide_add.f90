! vybe-test: fortran/arithmetic_extended/chained_subtract_divide_add
! origin: languages/fortran/tests/fortran/test_arithmetic_extended.rs
program t
if ((30 - 12 / 3 + 1) /= 27) then
    print *, "FAIL: want [27] got [", 30 - 12 / 3 + 1, "]"
    stop 1
end if
end program t
