! vybe-test: fortran/allocation_semantics/as_07
! origin: languages/fortran/tests/fortran/test_allocation_semantics.rs
program p
complex, allocatable :: a(:)
allocate(a(4))
end program p
