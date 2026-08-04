! vybe-test: fortran/expression_precedence/multiply_before_add_six_two_plus_one
! origin: languages/fortran/tests/fortran/test_expression_precedence.rs
program t
if ((6 * 2 + 1) /= 13) then
    print *, "FAIL: want [13] got [", 6 * 2 + 1, "]"
    stop 1
end if
end program t
