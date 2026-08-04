! vybe-test: fortran/complex/cmplx_one_arg
! origin: languages/fortran/tests/fortran/test_complex.rs
program t
  complex :: z
  z = cmplx(5.0)
  print *, z
end program t
