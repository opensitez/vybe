! vybe-test: fortran/allocation_semantics/as_18
! origin: languages/fortran/tests/fortran/test_allocation_semantics.rs
program p
integer, allocatable :: a(:)
allocate(a(3))
deallocate(a)
end program p
