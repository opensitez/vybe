! vybe-test: fortran/arithmetic_extended/real_minus_integer_literal
! origin: languages/fortran/tests/fortran/test_arithmetic_extended.rs
program t
if ((10.0 - 3) /= 7) then
    print *, "FAIL: want [7] got [", 10.0 - 3, "]"
    stop 1
end if
end program t
