! vybe-test: fortran/complex/conjg_basic
! origin: languages/fortran/tests/fortran/test_complex.rs
program t
  complex :: z = (3.0, 4.0)
  complex :: c
  c = conjg(z)
  print *, c
end program t
