! vybe-test: fortran/arithmetic_extended/integer_plus_real_literal
! origin: languages/fortran/tests/fortran/test_arithmetic_extended.rs
program t
if ((2 + 3.0) /= 5) then
    print *, "FAIL: want [5] got [", 2 + 3.0, "]"
    stop 1
end if
end program t
