! vybe-test: fortran/kinds/integer_kind_8
! origin: languages/fortran/tests/fortran/test_kinds.rs
program t
  integer(kind=8) :: big = 1000000000
  print *, big
end program t
