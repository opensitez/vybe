! vybe-test: fortran/variable_declarations_extended/double_precision_dimension_array
! origin: languages/fortran/tests/fortran/test_variable_declarations_extended.rs
program t
double precision, dimension(2) :: vals
vals(1) = 1.5d0
vals(2) = 2.5d0
if (abs((vals(2)) - 2.5) > 1.0e-6) then
    print *, "FAIL: want [2.5] got [", vals(2), "]"
    stop 1
end if
end program t
