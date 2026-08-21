! vybe-test: fortran/ieee/ieee_support_halting
! origin: languages/fortran/tests/fortran/test_fortran2018.rs

program test
    use ieee_arithmetic
    print *, ieee_support_halting(ieee_divide_by_zero)
end program test
