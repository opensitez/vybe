! vybe-test: fortran/pointers/pointer_real
! origin: languages/fortran/tests/fortran/test_pointers.rs
program t
  real, pointer :: p => null()
  print *, associated(p)
end program t
