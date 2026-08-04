! vybe-test: fortran/allocation_semantics/as_15
! origin: languages/fortran/tests/fortran/test_allocation_semantics.rs
program p
integer, allocatable :: a(:,:,:)
allocate(a(2,2,2))
end program p
