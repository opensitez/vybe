! vybe-test: fortran/math_f2008/norm2_double
! origin: languages/fortran/tests/fortran/test_math_f2008.rs

program test
    real(kind=8) :: v(4) = [1.0d0, 1.0d0, 1.0d0, 1.0d0]
    if (abs((norm2(v)) - (2.0)) > 2.000000e-05) then
        print *, "FAIL: want [2.0] got [", norm2(v), "]"
        stop 1
    end if
end program test
