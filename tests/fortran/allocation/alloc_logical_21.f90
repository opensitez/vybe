! vybe-test: fortran/allocation/alloc_logical_21
! origin: languages/fortran/tests/fortran/test_allocation.rs
program p
logical, allocatable :: a(:)
allocate(a(2))
end program p
