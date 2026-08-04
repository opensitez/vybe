! vybe-test: fortran/complex/real_part
! origin: languages/fortran/tests/fortran/test_complex.rs
program t
  complex :: z = (3.0, 4.0)
  print *, real(z)
end program t
