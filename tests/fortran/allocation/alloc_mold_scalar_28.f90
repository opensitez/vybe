! vybe-test: fortran/allocation/alloc_mold_scalar_28
! origin: languages/fortran/tests/fortran/test_allocation.rs
program p
integer, allocatable :: a,b
allocate(b)
allocate(a, mold=b)
end program p
