! vybe-test: fortran/math_f2008/erf_erfc_sum_is_one
! origin: languages/fortran/tests/fortran/test_math_f2008.rs

program test
    real :: x = 1.5
    real :: total
    total = erf(x) + erfc(x)
    print *, total
end program test
