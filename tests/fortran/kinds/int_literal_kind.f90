! vybe-test: fortran/kinds/int_literal_kind
! origin: languages/fortran/tests/fortran/test_kinds.rs
program t
  integer :: x = 100_4
  print *, x
end program t
