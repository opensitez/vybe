! vybe-test: fortran/allocation_semantics/as_09
! origin: languages/fortran/tests/fortran/test_allocation_semantics.rs
program p
character(len=:), allocatable :: s
allocate(character(len=5) :: s)
end program p
