! vybe-test: fortran/complex/abs_unit_imaginary
! origin: languages/fortran/tests/fortran/test_complex.rs
program t
  complex :: z = (0.0, 1.0)
  print *, abs(z)
end program t
