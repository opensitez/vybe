! vybe-test: fortran/kinds/real_kind_16
! origin: languages/fortran/tests/fortran/test_kinds.rs
program t
  real(kind=16) :: q = 3.14159265358979_16
  print *, q
end program t
