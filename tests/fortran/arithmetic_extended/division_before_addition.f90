! vybe-test: fortran/arithmetic_extended/division_before_addition
! origin: languages/fortran/tests/fortran/test_arithmetic_extended.rs
program t
if ((20 / 4 + 3) /= 8) then
    print *, "FAIL: want [8] got [", 20 / 4 + 3, "]"
    stop 1
end if
end program t
