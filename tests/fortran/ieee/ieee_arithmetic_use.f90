! vybe-test: fortran/ieee/ieee_arithmetic_use
! origin: languages/fortran/tests/fortran/test_fortran2008.rs

program test
    use ieee_arithmetic
    real :: x
    x = ieee_value(x, ieee_positive_inf)
    print *, ieee_is_finite(x)
end program test
