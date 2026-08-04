! vybe-test: fortran/math_f2008/norm2_double
! origin: languages/fortran/tests/fortran/test_math_f2008.rs

program test
    real(kind=8) :: v(4) = [1.0d0, 1.0d0, 1.0d0, 1.0d0]
    print *, norm2(v)
end program test
