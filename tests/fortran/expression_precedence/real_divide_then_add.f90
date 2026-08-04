! vybe-test: fortran/expression_precedence/real_divide_then_add
! origin: languages/fortran/tests/fortran/test_expression_precedence.rs
program t
if ((4 / 2 + 1.0) /= 3) then
    print *, "FAIL: want [3] got [", 4 / 2 + 1.0, "]"
    stop 1
end if
end program t
