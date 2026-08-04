! vybe-test: fortran/expression_precedence/power_right_assoc_five_squared_cubed
! origin: languages/fortran/tests/fortran/test_expression_precedence.rs
program t
if ((5 ** 2 ** 3) /= 390625) then
    print *, "FAIL: want [390625] got [", 5 ** 2 ** 3, "]"
    stop 1
end if
end program t
