! vybe-test: fortran/arithmetic_extended/real_division_negative_over_positive
! origin: languages/fortran/tests/fortran/test_arithmetic_extended.rs
program t
if ((-8.0 / 4.0) /= -2) then
    print *, "FAIL: want [-2] got [", -8.0 / 4.0, "]"
    stop 1
end if
end program t
