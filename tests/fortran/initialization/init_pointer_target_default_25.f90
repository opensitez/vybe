! vybe-test: fortran/initialization/init_pointer_target_default_25
! origin: languages/fortran/tests/fortran/test_initialization.rs
program driver
integer, target :: t = 7
integer, pointer :: p => t
print *, p
end program driver