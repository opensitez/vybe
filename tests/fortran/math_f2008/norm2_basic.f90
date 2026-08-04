! vybe-test: fortran/math_f2008/norm2_basic
! origin: languages/fortran/tests/fortran/test_math_f2008.rs

program test
    real :: a(3) = [3.0, 4.0, 0.0]
    print *, norm2(a)
end program test
