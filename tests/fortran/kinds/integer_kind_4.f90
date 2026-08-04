! vybe-test: fortran/kinds/integer_kind_4
! origin: languages/fortran/tests/fortran/test_kinds.rs
program t
  integer(kind=4) :: x = 100
  print *, x
end program t
