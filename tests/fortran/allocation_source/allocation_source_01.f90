! vybe-test: fortran/allocation_source/allocation_source_01
! origin: languages/fortran/tests/fortran/test_allocation_source.rs
program p
integer, allocatable :: a(:)
allocate(a(3), source=[1,2,3])
end program p
