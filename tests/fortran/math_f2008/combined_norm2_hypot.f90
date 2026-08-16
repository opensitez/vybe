! vybe-test: fortran/math_f2008/combined_norm2_hypot
! origin: languages/fortran/tests/fortran/test_math_f2008.rs

program test
    real :: v(2) = [3.0, 4.0]
    real :: h
    h = hypot(v(1), v(2))
    if (abs((norm2(v)) - (5.0)) > 5.000000e-05) then
        print *, "FAIL: want [5.0] got [", norm2(v), "]"
        stop 1
    end if
    if (abs((h) - (5.0)) > 5.000000e-05) then
        print *, "FAIL: want [5.0] got [", h, "]"
        stop 1
    end if
end program test
