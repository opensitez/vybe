! vybe-test: fortran/expression_precedence/power_right_assoc_two_cubed_squared
! origin: languages/fortran/tests/fortran/test_expression_precedence.rs
program t
if ((2 ** 3 ** 2) /= 512) then
    print *, "FAIL: want [512] got [", 2 ** 3 ** 2, "]"
    stop 1
end if
end program t
