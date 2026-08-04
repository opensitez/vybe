! vybe-test: fortran/arithmetic_extended/chained_unary_in_expression
! origin: languages/fortran/tests/fortran/test_arithmetic_extended.rs
program t
integer :: a = 3, b = 7
if ((a + -b) /= -4) then
    print *, "FAIL: want [-4] got [", a + -b, "]"
    stop 1
end if
end program t
