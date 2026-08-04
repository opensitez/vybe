! vybe-test: fortran/expression_precedence/multiply_before_add_three_four_five
! origin: languages/fortran/tests/fortran/test_expression_precedence.rs
program t
if ((3 + 4 * 5) /= 23) then
    print *, "FAIL: want [23] got [", 3 + 4 * 5, "]"
    stop 1
end if
end program t
