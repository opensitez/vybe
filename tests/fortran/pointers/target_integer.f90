! vybe-test: fortran/pointers/target_integer
! origin: languages/fortran/tests/fortran/test_pointers.rs
program t
  integer, target :: x = 42
  integer, pointer :: p
  p => x
  print *, p
end program t
