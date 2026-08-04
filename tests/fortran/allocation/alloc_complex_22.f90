! vybe-test: fortran/allocation/alloc_complex_22
! origin: languages/fortran/tests/fortran/test_allocation.rs
program p
complex, allocatable :: a(:)
allocate(a(2))
end program p
