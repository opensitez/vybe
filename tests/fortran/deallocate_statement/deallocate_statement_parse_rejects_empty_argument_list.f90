! vybe-test: fortran/deallocate_statement/deallocate_statement_parse_rejects_empty_argument_list
! origin: languages/fortran/tests/fortran/test_deallocate_statement.rs
program t
integer, allocatable :: a(:)
deallocate()
end program t
