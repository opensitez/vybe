! vybe-test: fortran/math_f2008/atanh_tanh_roundtrip
! origin: languages/fortran/tests/fortran/test_math_f2008.rs

program test
    real :: x = 0.5
    if (abs((atanh(tanh(x))) - (0.5)) > 5.000000e-06) then
        print *, "FAIL: want [0.5] got [", atanh(tanh(x)), "]"
        stop 1
    end if
end program test
