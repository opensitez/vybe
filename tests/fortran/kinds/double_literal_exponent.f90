! vybe-test: fortran/kinds/double_literal_exponent
! origin: languages/fortran/tests/fortran/test_kinds.rs
program t
  real(kind=8) :: d = 1.23d+10
  print *, d
end program t
