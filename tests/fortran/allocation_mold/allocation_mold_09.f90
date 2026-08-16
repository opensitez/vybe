! vybe-test: fortran/allocation_mold/allocation_mold_09
! origin: languages/fortran/tests/fortran/test_allocation_mold.rs
program driver
integer, pointer :: p(:)
integer, allocatable :: a(:)
allocate(p(3))
allocate(a, mold=p)
end program driver