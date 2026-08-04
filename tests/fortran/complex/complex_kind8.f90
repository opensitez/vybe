! vybe-test: fortran/complex/complex_kind8
! origin: languages/fortran/tests/fortran/test_complex.rs
program t
  complex(kind=8) :: z = (1.0_8, 2.0_8)
  print *, z
end program t
