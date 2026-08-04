! vybe-test: fortran/allocation_semantics/as_10
! origin: languages/fortran/tests/fortran/test_allocation_semantics.rs
program p
integer, allocatable :: x
allocate(x, source=5)
end program p
