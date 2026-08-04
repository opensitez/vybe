! vybe-test: fortran/pointers/target_array
! origin: languages/fortran/tests/fortran/test_pointers.rs
program t
  integer, target :: a(5) = [1,2,3,4,5]
  integer, pointer :: p(:)
  p => a
  print *, p(1)
end program t
