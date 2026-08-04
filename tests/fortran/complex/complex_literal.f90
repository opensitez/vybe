! vybe-test: fortran/complex/complex_literal
! origin: languages/fortran/tests/fortran/test_complex.rs
program t
  complex :: z = (3.0, 4.0)
  print *, z
end program t
