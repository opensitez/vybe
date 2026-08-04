! vybe-test: fortran/allocation_source/allocation_source_02
! origin: languages/fortran/tests/fortran/test_allocation_source.rs
program p
integer, allocatable :: x
allocate(x, source=5)
end program p
