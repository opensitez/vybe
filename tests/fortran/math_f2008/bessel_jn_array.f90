! vybe-test: fortran/math_f2008/bessel_jn_array
! origin: languages/fortran/tests/fortran/test_math_f2008.rs

program test
    real :: values(3)
    values = bessel_jn(0, 2, 1.0)
    if (abs((values(1)) - (0.765197635)) > 7.651976e-06) then
        print *, "FAIL: want [0.765197635] got [", values(1), "]"
        stop 1
    end if
end program test
