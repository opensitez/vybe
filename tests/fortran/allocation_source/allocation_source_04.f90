! vybe-test: fortran/allocation_source/allocation_source_04
! origin: languages/fortran/tests/fortran/test_allocation_source.rs
program p
character(len=:), allocatable :: s
allocate(character(len=3) :: s, source='abc')
end program p
