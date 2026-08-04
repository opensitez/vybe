! vybe-test: fortran/complex/cos_complex
! origin: languages/fortran/tests/fortran/test_complex.rs
program t
  complex :: z = (0.0, 0.0)
  complex :: r
  r = cos(z)
  print *, r
end program t
