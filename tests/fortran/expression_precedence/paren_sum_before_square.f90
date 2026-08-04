! vybe-test: fortran/expression_precedence/paren_sum_before_square
! origin: languages/fortran/tests/fortran/test_expression_precedence.rs
program t
if (((1 + 2) ** 2) /= 9) then
    print *, "FAIL: want [9] got [", (1 + 2) ** 2, "]"
    stop 1
end if
end program t
