! vybe-test: fortran/math_f2008/inverse_hyp_double
! origin: languages/fortran/tests/fortran/test_math_f2008.rs

program test
    real(kind=8) :: x = 1.0d0
    print *, acosh(cosh(x))
    print *, asinh(sinh(x))
    print *, atanh(tanh(x * 0.5d0))
end program test
