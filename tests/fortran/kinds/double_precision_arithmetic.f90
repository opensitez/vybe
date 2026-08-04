! vybe-test: fortran/kinds/double_precision_arithmetic
! origin: languages/fortran/tests/fortran/test_kinds.rs
program t
  double precision :: a = 1.0, b = 3.0, c
  c = a / b
  print *, c
end program t
