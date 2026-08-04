! vybe-test: fortran/expression_precedence/int_real_mult_in_add_chain
! origin: languages/fortran/tests/fortran/test_expression_precedence.rs
program t
if ((1 + 2 * 3.0) /= 7) then
    print *, "FAIL: want [7] got [", 1 + 2 * 3.0, "]"
    stop 1
end if
end program t
