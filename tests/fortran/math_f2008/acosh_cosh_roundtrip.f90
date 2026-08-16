! vybe-test: fortran/math_f2008/acosh_cosh_roundtrip
! origin: languages/fortran/tests/fortran/test_math_f2008.rs

program test
    real :: x = 2.0
    if (abs((acosh(cosh(x))) - (2.0)) > 2.000000e-05) then
        print *, "FAIL: want [2.0] got [", acosh(cosh(x)), "]"
        stop 1
    end if
end program test
