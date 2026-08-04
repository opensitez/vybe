! vybe-test: fortran/deallocate_statement/deallocate_statement_04
! origin: languages/fortran/tests/fortran/test_deallocate_statement.rs
program p
integer, allocatable :: x
allocate(x)
deallocate(x)
end program p
