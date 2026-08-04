! vybe-test: fortran/complex/aimag_part
! origin: languages/fortran/tests/fortran/test_complex.rs
program t
  complex :: z = (3.0, 4.0)
  print *, aimag(z)
end program t
