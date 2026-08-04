! vybe-test: fortran/math_f2008/combined_norm2_hypot
! origin: languages/fortran/tests/fortran/test_math_f2008.rs

program test
    real :: v(2) = [3.0, 4.0]
    real :: h
    h = hypot(v(1), v(2))
    print *, norm2(v)
    print *, h
end program test
