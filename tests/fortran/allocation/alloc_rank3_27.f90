! vybe-test: fortran/allocation/alloc_rank3_27
! origin: languages/fortran/tests/fortran/test_allocation.rs
program p
integer, allocatable :: a(:,:,:)
allocate(a(2,2,2))
end program p
