! vybe-test: fortran/out_of_range/out_of_range_real_infinity
! origin: languages/fortran/tests/fortran/test_fortran2018.rs

program test
    use ieee_arithmetic
    real :: x
    x = ieee_value(x, ieee_positive_inf)
    print *, out_of_range(x, 0)
end program test
