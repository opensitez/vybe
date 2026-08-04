! vybe-test: fortran/expression_precedence/power_then_real_add
! origin: languages/fortran/tests/fortran/test_expression_precedence.rs
program t
if ((2 ** 2 + 1.0) /= 5) then
    print *, "FAIL: want [5] got [", 2 ** 2 + 1.0, "]"
    stop 1
end if
end program t
