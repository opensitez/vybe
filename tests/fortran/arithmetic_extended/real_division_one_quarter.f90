! vybe-test: fortran/arithmetic_extended/real_division_one_quarter
! origin: languages/fortran/tests/fortran/test_arithmetic_extended.rs
program t
if (abs((1.0 / 4.0) - 0.25) > 1.0e-6) then
    print *, "FAIL: want [0.25] got [", 1.0 / 4.0, "]"
    stop 1
end if
end program t
