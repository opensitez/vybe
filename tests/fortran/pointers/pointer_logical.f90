! vybe-test: fortran/pointers/pointer_logical
! origin: languages/fortran/tests/fortran/test_pointers.rs
program t
  logical, pointer :: p => null()
  print *, associated(p)
end program t
