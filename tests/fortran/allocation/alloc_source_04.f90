! vybe-test: fortran/allocation/alloc_source_04
! origin: languages/fortran/tests/fortran/test_allocation.rs
program p
integer, allocatable :: a(:)
allocate(a(3), source=[1,2,3])
end program p
