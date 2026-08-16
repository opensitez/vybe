! vybe-test: fortran/math_f2008/erf_erfc_sum_is_one
! origin: languages/fortran/tests/fortran/test_math_f2008.rs

program test
    real :: x = 1.5
    real :: total
    total = erf(x) + erfc(x)
    if (abs((total) - (1.0)) > 1.000000e-05) then
        print *, "FAIL: want [1.0] got [", total, "]"
        stop 1
    end if
end program test
