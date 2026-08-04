! vybe-test: fortran/pointers/pointer_char
! origin: languages/fortran/tests/fortran/test_pointers.rs
program t
  character(len=10), pointer :: p => null()
  print *, associated(p)
end program t
