! vybe-test: fortran/variables/parameter_real
! origin: languages/fortran/tests/fortran/test_variables.rs
program t
real, parameter :: PI = 3.14159
if (abs((PI) - 3.14159) > 1.0e-6) then
    print *, "FAIL: want [3.14159] got [", PI, "]"
    stop 1
end if
end program t
