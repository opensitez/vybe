! vybe-test: fortran/kinds/double_precision_parameter
! origin: languages/fortran/tests/fortran/test_kinds.rs
program t
  double precision, parameter :: PI = 3.141592653589793
  print *, PI
end program t
