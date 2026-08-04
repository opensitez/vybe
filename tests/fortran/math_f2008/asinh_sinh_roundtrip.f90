! vybe-test: fortran/math_f2008/asinh_sinh_roundtrip
! origin: languages/fortran/tests/fortran/test_math_f2008.rs

program test
    real :: x = 1.5
    print *, asinh(sinh(x))
end program test
