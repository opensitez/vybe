! vybe-test: fortran/allocation/deferred_len_char_14
! origin: languages/fortran/tests/fortran/test_allocation.rs
program p
character(len=:), allocatable :: s
allocate(character(len=3) :: s)
end program p
