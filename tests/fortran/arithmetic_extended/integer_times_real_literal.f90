! vybe-test: fortran/arithmetic_extended/integer_times_real_literal
! origin: languages/fortran/tests/fortran/test_arithmetic_extended.rs
program t
if ((5 * 2.0) /= 10) then
    print *, "FAIL: want [10] got [", 5 * 2.0, "]"
    stop 1
end if
end program t
