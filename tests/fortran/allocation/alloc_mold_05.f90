! vybe-test: fortran/allocation/alloc_mold_05
! origin: languages/fortran/tests/fortran/test_allocation.rs
program p
integer, allocatable :: a(:), b(:)
allocate(b(3))
allocate(a, mold=b)
end program p
