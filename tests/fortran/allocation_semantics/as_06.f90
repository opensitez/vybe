! vybe-test: fortran/allocation_semantics/as_06
! origin: languages/fortran/tests/fortran/test_allocation_semantics.rs
program p
real, allocatable :: a(:,:)
allocate(a(2,3))
end program p
