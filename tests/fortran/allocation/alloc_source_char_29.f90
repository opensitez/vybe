! vybe-test: fortran/allocation/alloc_source_char_29
! origin: languages/fortran/tests/fortran/test_allocation.rs
program p
character(len=:), allocatable :: s
allocate(character(len=4) :: s)
s='test'
end program p
