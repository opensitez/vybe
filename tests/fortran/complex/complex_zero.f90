! vybe-test: fortran/complex/complex_zero
! origin: languages/fortran/tests/fortran/test_complex.rs
program t
  complex :: z = (0.0, 0.0)
  print *, z
end program t
