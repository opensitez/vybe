! vybe-test: fortran/math_f2008/inverse_hyp_double
! origin: languages/fortran/tests/fortran/test_math_f2008.rs

program test
    real(kind=8) :: x = 1.0d0
    if (abs((acosh(cosh(x))) - (1.0)) > 1.000000e-05) then
        print *, "FAIL: want [1.0] got [", acosh(cosh(x)), "]"
        stop 1
    end if
    if (abs((asinh(sinh(x))) - (1.0)) > 1.000000e-05) then
        print *, "FAIL: want [1.0] got [", asinh(sinh(x)), "]"
        stop 1
    end if
    if (abs((atanh(tanh(x * 0.5d0))) - (0.5)) > 5.000000e-06) then
        print *, "FAIL: want [0.5] got [", atanh(tanh(x * 0.5d0)), "]"
        stop 1
    end if
end program test
