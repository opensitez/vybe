! vybe-test: fortran/complex/cmplx_from_int
! origin: languages/fortran/tests/fortran/test_complex.rs
program t
  complex :: z
  integer :: a = 3, b = 4
  z = cmplx(a, b)
  print *, z
end program t
