! vybe-test: fortran/allocation_mold/allocation_mold_10
! origin: languages/fortran/tests/fortran/test_allocation_mold.rs
program p
real, allocatable :: a(:,:,:), b(:,:,:)
allocate(b(2,2,2))
allocate(a, mold=b)
end program p
