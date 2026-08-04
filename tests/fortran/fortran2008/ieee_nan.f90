! vybe-test: fortran/fortran2008/ieee_nan
! origin: languages/fortran/tests/fortran/test_fortran2008.rs

program test
    use ieee_arithmetic
    real :: x
    x = ieee_value(x, ieee_quiet_nan)
    print *, ieee_is_nan(x)
end program test
