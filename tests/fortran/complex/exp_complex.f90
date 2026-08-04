! vybe-test: fortran/complex/exp_complex
! origin: languages/fortran/tests/fortran/test_complex.rs
program t
  complex :: z = (0.0, 3.14159)
  complex :: r
  r = exp(z)
  print *, r
end program t
