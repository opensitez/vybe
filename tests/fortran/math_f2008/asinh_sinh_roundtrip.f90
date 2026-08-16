! vybe-test: fortran/math_f2008/asinh_sinh_roundtrip
! origin: languages/fortran/tests/fortran/test_math_f2008.rs

program test
    real :: x = 1.5
    if (abs((asinh(sinh(x))) - (1.5)) > 1.500000e-05) then
        print *, "FAIL: want [1.5] got [", asinh(sinh(x)), "]"
        stop 1
    end if
end program test
