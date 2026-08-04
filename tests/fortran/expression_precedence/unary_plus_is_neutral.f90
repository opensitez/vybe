! vybe-test: fortran/expression_precedence/unary_plus_is_neutral
! origin: languages/fortran/tests/fortran/test_expression_precedence.rs
program t
if ((+(-2)) /= -2) then
    print *, "FAIL: want [-2] got [", +(-2), "]"
    stop 1
end if
end program t
