! vybe-test: fortran/kinds/double_literal
! origin: languages/fortran/tests/fortran/test_kinds.rs
program t
  real(kind=8) :: d = 1.0d0
  print *, d
end program t
