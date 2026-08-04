! vybe-test: fortran/pointers/pointer_integer
! origin: languages/fortran/tests/fortran/test_pointers.rs
program t
  integer, pointer :: p => null()
  print *, associated(p)
end program t
