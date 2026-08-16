! vybe-test: fortran/math_f2008/erf_with_gamma
! origin: languages/fortran/tests/fortran/test_math_f2008.rs

program test
    real :: x = 0.5
    if (abs((erf(x)) - (0.520499885)) > 5.204999e-06) then
        print *, "FAIL: want [0.520499885] got [", erf(x), "]"
        stop 1
    end if
    if (abs((gamma(x + 0.5) / gamma(0.5)) - (0.564189553)) > 5.641896e-06) then
        print *, "FAIL: want [0.564189553] got [", gamma(x + 0.5) / gamma(0.5), "]"
        stop 1
    end if
end program test
