! vybe-test: fortran/complex/sqrt_complex
! origin: languages/fortran/tests/fortran/test_complex.rs
program t
  complex :: z = (-1.0, 0.0)
  complex :: r
  r = sqrt(z)
  print *, r
end program t
