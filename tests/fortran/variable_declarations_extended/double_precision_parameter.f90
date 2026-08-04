! vybe-test: fortran/variable_declarations_extended/double_precision_parameter
! origin: languages/fortran/tests/fortran/test_variable_declarations_extended.rs
program t
double precision, parameter :: e = 2.718281828d0
if (abs((e) - 2.718281828) > 1.0e-6) then
    print *, "FAIL: want [2.718281828] got [", e, "]"
    stop 1
end if
end program t
