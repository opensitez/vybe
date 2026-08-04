! vybe-test: fortran/expression_precedence/paren_quotient_of_differences
! origin: languages/fortran/tests/fortran/test_expression_precedence.rs
program t
if (((10 - 2) / (3 - 1)) /= 4) then
    print *, "FAIL: want [4] got [", (10 - 2) / (3 - 1), "]"
    stop 1
end if
end program t
