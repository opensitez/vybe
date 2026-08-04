! vybe-test: fortran/expression_precedence/divide_multiply_left_assoc_eight
! origin: languages/fortran/tests/fortran/test_expression_precedence.rs
program t
if ((8 / 2 * 3) /= 12) then
    print *, "FAIL: want [12] got [", 8 / 2 * 3, "]"
    stop 1
end if
end program t
