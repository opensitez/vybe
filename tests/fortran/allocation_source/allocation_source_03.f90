! vybe-test: fortran/allocation_source/allocation_source_03
! origin: languages/fortran/tests/fortran/test_allocation_source.rs
program p
real, allocatable :: a(:,:)
allocate(a(2,2), source=reshape([1.0,2.0,3.0,4.0],[2,2]))
end program p
