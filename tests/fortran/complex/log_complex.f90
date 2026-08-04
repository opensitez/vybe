! vybe-test: fortran/complex/log_complex
! origin: languages/fortran/tests/fortran/test_complex.rs
program t
  complex :: z = (1.0, 0.0)
  complex :: r
  r = log(z)
  print *, r
end program t
