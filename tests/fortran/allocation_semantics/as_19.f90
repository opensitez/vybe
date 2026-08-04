! vybe-test: fortran/allocation_semantics/as_19
! origin: languages/fortran/tests/fortran/test_allocation_semantics.rs
program p
integer, allocatable :: a(:)
allocate(a(3))
a = [1,2,3]
end program p
