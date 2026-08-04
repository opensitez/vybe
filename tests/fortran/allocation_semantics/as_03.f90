! vybe-test: fortran/allocation_semantics/as_03
! origin: languages/fortran/tests/fortran/test_allocation_semantics.rs
program p
integer, allocatable :: a(:),b(:)
allocate(b(2))
allocate(a, mold=b)
end program p
