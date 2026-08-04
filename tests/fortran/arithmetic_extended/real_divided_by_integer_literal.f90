! vybe-test: fortran/arithmetic_extended/real_divided_by_integer_literal
! origin: languages/fortran/tests/fortran/test_arithmetic_extended.rs
program t
if ((6.0 / 2) /= 3) then
    print *, "FAIL: want [3] got [", 6.0 / 2, "]"
    stop 1
end if
end program t
