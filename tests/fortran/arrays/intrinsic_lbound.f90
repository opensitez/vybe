! vybe-test: fortran/arrays/intrinsic_lbound
! origin: languages/fortran/tests/fortran/test_arrays.rs
program t
  integer :: a(3)
  print *, lbound(a, 1)
end program t
