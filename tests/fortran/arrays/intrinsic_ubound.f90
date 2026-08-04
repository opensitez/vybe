! vybe-test: fortran/arrays/intrinsic_ubound
! origin: languages/fortran/tests/fortran/test_arrays.rs
program t
  integer :: a(3)
  print *, ubound(a, 1)
end program t
