! vybe-test: fortran/named_constant_statement_functions/test_named_constant_statement_functions_evaluate_with_constants
! origin: languages/fortran/tests/fortran/test_named_constant_statement_functions.rs
program test_named_constant_statement_functions
    integer :: n
    integer :: cube
    cube(n) = n ** 3
    n = 4
    if ((cube(3)) /= 27) then
    print *, "FAIL: want [27] got [", cube(3), "]"
    stop 1
end if
end program test_named_constant_statement_functions
