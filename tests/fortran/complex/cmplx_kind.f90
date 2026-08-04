! vybe-test: fortran/complex/cmplx_kind
! origin: languages/fortran/tests/fortran/test_complex.rs
program t
  complex(kind=8) :: z
  z = cmplx(1.0, 2.0, kind=8)
  print *, z
end program t
