! vybe-test: fortran/allocation/dealloc_scalar_16
! origin: languages/fortran/tests/fortran/test_allocation.rs
program p
integer, allocatable :: x
allocate(x)
deallocate(x)
end program p
