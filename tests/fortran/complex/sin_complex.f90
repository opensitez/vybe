! vybe-test: fortran/complex/sin_complex
! origin: languages/fortran/tests/fortran/test_complex.rs
program t
  complex :: z = (0.0, 0.0)
  complex :: r
  r = sin(z)
  print *, r
end program t
