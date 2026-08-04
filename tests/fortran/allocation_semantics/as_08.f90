! vybe-test: fortran/allocation_semantics/as_08
! origin: languages/fortran/tests/fortran/test_allocation_semantics.rs
program p
logical, allocatable :: a(:)
allocate(a(5))
end program p
