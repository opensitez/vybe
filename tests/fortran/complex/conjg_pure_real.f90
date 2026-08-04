! vybe-test: fortran/complex/conjg_pure_real
! origin: languages/fortran/tests/fortran/test_complex.rs
program t
  complex :: z = (5.0, 0.0)
  print *, conjg(z)
end program t
