! vybe-test: fortran/math_f2008/acosh_cosh_roundtrip
! origin: languages/fortran/tests/fortran/test_math_f2008.rs

program test
    real :: x = 2.0
    print *, acosh(cosh(x))
end program test
