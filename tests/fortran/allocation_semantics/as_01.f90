! vybe-test: fortran/allocation_semantics/as_01
! origin: languages/fortran/tests/fortran/test_allocation_semantics.rs
program p
integer, allocatable :: a(:)
allocate(a(3))
print *, size(a)
end program p
