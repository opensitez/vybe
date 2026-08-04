! vybe-test: fortran/allocate_statement/allocate_statement_parse_rejects_missing_array_shape
! origin: languages/fortran/tests/fortran/test_allocate_statement.rs
program p
integer, allocatable :: a(:)
allocate(a)
end program p
