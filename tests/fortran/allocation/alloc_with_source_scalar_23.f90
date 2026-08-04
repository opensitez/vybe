! vybe-test: fortran/allocation/alloc_with_source_scalar_23
! origin: languages/fortran/tests/fortran/test_allocation.rs
program p
integer, allocatable :: x
allocate(x, source=5)
end program p
