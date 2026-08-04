! vybe-test: fortran/variable_declarations_extended/double_precision_init_literal
! origin: languages/fortran/tests/fortran/test_variable_declarations_extended.rs
program t
double precision :: d = 9.87654321d0
if (abs((d) - 9.87654321) > 1.0e-6) then
    print *, "FAIL: want [9.87654321] got [", d, "]"
    stop 1
end if
end program t
