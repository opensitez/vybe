! vybe-test: fortran/kinds/double_precision_assign
! origin: languages/fortran/tests/fortran/test_kinds.rs
program t
  double precision :: d
  d = 3.141592653589793
  print *, d
end program t
