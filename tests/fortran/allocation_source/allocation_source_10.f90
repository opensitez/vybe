! vybe-test: fortran/allocation_source/allocation_source_10
! origin: languages/fortran/tests/fortran/test_allocation_source.rs
program p
integer, allocatable :: a(:,:,:)
allocate(a(2,2,1), source=reshape([1,2,3,4],[2,2,1]))
end program p
