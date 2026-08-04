! vybe-test: fortran/deallocate_statement/deallocate_statement_03
! origin: languages/fortran/tests/fortran/test_deallocate_statement.rs
program p
character(len=:), allocatable :: s
allocate(character(len=3) :: s)
deallocate(s)
end program p
