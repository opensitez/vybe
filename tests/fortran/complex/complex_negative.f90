! vybe-test: fortran/complex/complex_negative
! origin: languages/fortran/tests/fortran/test_complex.rs
program t
  complex :: z = (-1.0, -2.0)
  print *, z
end program t
