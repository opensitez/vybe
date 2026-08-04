! vybe-test: fortran/initialization/init_allocatable_default_26
! origin: languages/fortran/tests/fortran/test_initialization.rs
program p
integer, allocatable :: a(:)
a = [1,2,3]
print *, a(1)
end program p
