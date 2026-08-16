! vybe-test: fortran/parameter_expression_constant_folding/test_parameter_expression_constant_folding_precomputes_arithmetic
! origin: languages/fortran/tests/fortran/test_parameter_expression_constant_folding.rs

program test_parameter_expression_constant_folding
    integer, parameter :: a = 2 + 3 * 4 - 1
    integer, parameter :: b = merge(10, 1, a > 10)
    if ((a) /= 13) then
    print *, "FAIL: want [13] got [", a, "]"
    stop 1
end if
    if ((b) /= 10) then
    print *, "FAIL: want [10] got [", b, "]"
    stop 1
end if
end program test_parameter_expression_constant_folding
