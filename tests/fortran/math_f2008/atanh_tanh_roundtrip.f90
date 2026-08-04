! vybe-test: fortran/math_f2008/atanh_tanh_roundtrip
! origin: languages/fortran/tests/fortran/test_math_f2008.rs

program test
    real :: x = 0.5
    print *, atanh(tanh(x))
end program test
