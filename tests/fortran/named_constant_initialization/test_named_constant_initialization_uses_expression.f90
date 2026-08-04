! vybe-test: fortran/named_constant_initialization/test_named_constant_initialization_uses_expression
! origin: languages/fortran/tests/fortran/test_named_constant_initialization.rs

program test_named_constant_initialization
    real, parameter :: pi = 3.14159
    if ((nint(pi)) /= 3) then
    print *, "FAIL: want [3] got [", nint(pi), "]"
    stop 1
end if
end program test_named_constant_initialization
