! vybe-test: fortran/allocation_mold/allocation_mold_02
! origin: languages/fortran/tests/fortran/test_allocation_mold.rs
program p
integer, allocatable :: a,b
allocate(b)
allocate(a, mold=b)
end program p
