! vybe-test: fortran/allocate_statement/allocate_statement_03
! origin: languages/fortran/tests/fortran/test_allocate_statement.rs
program p
character(len=:), allocatable :: s
allocate(character(len=3) :: s)
end program p
