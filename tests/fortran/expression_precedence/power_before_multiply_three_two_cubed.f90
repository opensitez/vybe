! vybe-test: fortran/expression_precedence/power_before_multiply_three_two_cubed
! origin: languages/fortran/tests/fortran/test_expression_precedence.rs
program t
if ((3 * 2 ** 3) /= 24) then
    print *, "FAIL: want [24] got [", 3 * 2 ** 3, "]"
    stop 1
end if
end program t
