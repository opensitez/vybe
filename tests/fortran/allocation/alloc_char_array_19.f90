! vybe-test: fortran/allocation/alloc_char_array_19
! origin: languages/fortran/tests/fortran/test_allocation.rs
program p
character(len=4), allocatable :: a(:)
allocate(a(2))
end program p
