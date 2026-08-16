! vybe-test: fortran/math_f2008/log_gamma_vs_log_gamma
! origin: languages/fortran/tests/fortran/test_math_f2008.rs

program test
    real :: x = 10.0
    if (abs((log_gamma(x)) - (12.8018274)) > 1.280183e-04) then
        print *, "FAIL: want [12.8018274] got [", log_gamma(x), "]"
        stop 1
    end if
    if (abs((log(gamma(x))) - (12.8018274)) > 1.280183e-04) then
        print *, "FAIL: want [12.8018274] got [", log(gamma(x)), "]"
        stop 1
    end if
end program test
