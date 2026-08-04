! vybe-test: fortran/math_f2008/bessel_jn_array
! origin: languages/fortran/tests/fortran/test_math_f2008.rs

program test
    real :: values(3)
    values = bessel_jn(0, 2, 1.0)
    print *, values(1)
end program test
