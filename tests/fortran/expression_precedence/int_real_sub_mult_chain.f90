! vybe-test: fortran/expression_precedence/int_real_sub_mult_chain
! origin: languages/fortran/tests/fortran/test_expression_precedence.rs
program t
if ((10 - 3 * 2.0) /= 4) then
    print *, "FAIL: want [4] got [", 10 - 3 * 2.0, "]"
    stop 1
end if
end program t
