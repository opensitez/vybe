! vybe-test: fortran/kinds/real_literal_kind
! origin: languages/fortran/tests/fortran/test_kinds.rs
program t
  real :: x = 3.14_4
  print *, x
end program t
