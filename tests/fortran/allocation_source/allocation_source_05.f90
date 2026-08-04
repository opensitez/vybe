! vybe-test: fortran/allocation_source/allocation_source_05
! origin: languages/fortran/tests/fortran/test_allocation_source.rs
program p
character(len=:), allocatable :: a(:)
allocate(character(len=2) :: a(2), source=['aa','bb'])
end program p
