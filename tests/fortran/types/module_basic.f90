! vybe-test: fortran/types/module_basic
! origin: languages/fortran/tests/fortran/test_types.rs

module constants
    real, parameter :: PI = 3.14159
    real, parameter :: E = 2.71828
end module constants

program test
    use constants
    if (abs((PI) - 3.14159) > 1.0e-6) then
    print *, "FAIL: want [3.14159] got [", PI, "]"
    stop 1
end if
end program test
