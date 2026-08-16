! vybe-test: fortran/math_f2008/bessel_yn_array
! origin: languages/fortran/tests/fortran/test_math_f2008.rs

program test
    real :: values(3)
    values = bessel_yn(0, 2, 1.0)
    if (abs((values(1)) - (0.0882569626)) > 1.000000e-06) then
        print *, "FAIL: want [0.0882569626] got [", values(1), "]"
        stop 1
    end if
end program test
