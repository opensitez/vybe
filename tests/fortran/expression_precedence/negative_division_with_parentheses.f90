! vybe-test: fortran/expression_precedence/negative_division_with_parentheses
! origin: languages/fortran/tests/fortran/test_expression_precedence.rs
program t
if ((-(7 / 2)) /= -3) then
    print *, "FAIL: want [-3] got [", -(7 / 2), "]"
    stop 1
end if
if (((7 / 2)) /= 3) then
    print *, "FAIL: want [3] got [", (7 / 2), "]"
    stop 1
end if
end program t
