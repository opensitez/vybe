! vybe-test: fortran/kinds/integer_kind_2
! origin: languages/fortran/tests/fortran/test_kinds.rs
program t
  integer(kind=2) :: s = 32000
  print *, s
end program t
