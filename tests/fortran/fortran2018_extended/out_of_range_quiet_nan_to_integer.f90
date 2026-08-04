! vybe-test: fortran/fortran2018_extended/out_of_range_quiet_nan_to_integer
! origin: languages/fortran/tests/fortran/test_fortran2018_extended.rs

program t
    use ieee_arithmetic
    real :: x
    x = ieee_value(x, ieee_quiet_nan)
    print *, out_of_range(x, 0)
end program t
