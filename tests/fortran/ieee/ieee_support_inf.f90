! vybe-test: fortran/ieee/ieee_support_inf
! origin: languages/fortran/tests/fortran/test_fortran2018.rs

program test
    use ieee_arithmetic
    real :: x
    print *, ieee_support_inf(x)
end program test
