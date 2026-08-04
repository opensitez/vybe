! vybe-test: fortran/allocation/alloc_real_2d_20
! origin: languages/fortran/tests/fortran/test_allocation.rs
program p
real, allocatable :: a(:,:)
allocate(a(3,4))
end program p
