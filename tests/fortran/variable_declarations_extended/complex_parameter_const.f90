! vybe-test: fortran/variable_declarations_extended/complex_parameter_const
! origin: languages/fortran/tests/fortran/test_variable_declarations_extended.rs
program t
implicit none
complex, parameter :: unit_i = (0.0, 1.0)
if ((nint(aimag(unit_i))) /= 1) then
    print *, "FAIL: want [1] got [", nint(aimag(unit_i)), "]"
    stop 1
end if
end program t
