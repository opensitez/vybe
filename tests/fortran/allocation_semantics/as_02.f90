! vybe-test: fortran/allocation_semantics/as_02
! origin: languages/fortran/tests/fortran/test_allocation_semantics.rs
program p
integer, allocatable :: a(:)
allocate(a(3), source=[1,2,3])
print *, a
end program p
