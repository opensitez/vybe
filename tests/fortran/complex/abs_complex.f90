! vybe-test: fortran/complex/abs_complex
! origin: languages/fortran/tests/fortran/test_complex.rs
program t
  complex :: z = (3.0, 4.0)
  real :: m
  m = abs(z)
  print *, m
end program t
