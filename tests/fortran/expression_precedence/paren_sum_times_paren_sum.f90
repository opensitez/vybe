! vybe-test: fortran/expression_precedence/paren_sum_times_paren_sum
! origin: languages/fortran/tests/fortran/test_expression_precedence.rs
program t
if (((2 + 3) * (4 + 1)) /= 25) then
    print *, "FAIL: want [25] got [", (2 + 3) * (4 + 1), "]"
    stop 1
end if
end program t
