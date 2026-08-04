! vybe-test: fortran/allocation/allocate_scalar_15
! origin: languages/fortran/tests/fortran/test_allocation.rs
program p
integer, allocatable :: x
allocate(x)
end program p
