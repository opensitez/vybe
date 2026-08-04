! vybe-test: fortran/arithmetic_extended/unary_minus_on_literal
! origin: languages/fortran/tests/fortran/test_arithmetic_extended.rs
program t
if ((-17) /= -17) then
    print *, "FAIL: want [-17] got [", -17, "]"
    stop 1
end if
end program t
