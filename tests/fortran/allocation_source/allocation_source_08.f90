! vybe-test: fortran/allocation_source/allocation_source_08
! origin: languages/fortran/tests/fortran/test_allocation_source.rs
program p
complex, allocatable :: z(:)
allocate(z(2), source=[(1.0,2.0),(3.0,4.0)])
end program p
