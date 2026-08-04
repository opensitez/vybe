! vybe-test: fortran/pointers/target_real
! origin: languages/fortran/tests/fortran/test_pointers.rs
program t
  real, target :: x = 3.14
  real, pointer :: p
  p => x
  print *, p
end program t
