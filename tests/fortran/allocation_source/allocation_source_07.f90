! vybe-test: fortran/allocation_source/allocation_source_07
! origin: languages/fortran/tests/fortran/test_allocation_source.rs
program p
logical, allocatable :: l(:)
allocate(l(2), source=[.true.,.false.])
end program p
