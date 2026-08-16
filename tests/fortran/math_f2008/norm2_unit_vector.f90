! vybe-test: fortran/math_f2008/norm2_unit_vector
! origin: languages/fortran/tests/fortran/test_math_f2008.rs

program test
    real :: a(3) = [1.0, 0.0, 0.0]
    if (abs((norm2(a)) - (1.0)) > 1.000000e-05) then
        print *, "FAIL: want [1.0] got [", norm2(a), "]"
        stop 1
    end if
end program test
