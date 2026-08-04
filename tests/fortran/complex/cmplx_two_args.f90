! vybe-test: fortran/complex/cmplx_two_args
! origin: languages/fortran/tests/fortran/test_complex.rs
program t
  complex :: z
  z = cmplx(3.0, 4.0)
  print *, z
end program t
