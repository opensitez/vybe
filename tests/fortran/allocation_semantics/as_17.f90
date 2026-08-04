! vybe-test: fortran/allocation_semantics/as_17
! origin: languages/fortran/tests/fortran/test_allocation_semantics.rs
program p
integer, allocatable :: a(:)
allocate(a(0))
print *, size(a)
end program p
