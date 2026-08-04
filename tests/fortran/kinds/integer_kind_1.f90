! vybe-test: fortran/kinds/integer_kind_1
! origin: languages/fortran/tests/fortran/test_kinds.rs
program t
  integer(kind=1) :: b = 127
  print *, b
end program t
