! vybe-test: fortran/allocate_statement/allocate_statement_06
! origin: languages/fortran/tests/fortran/test_allocate_statement.rs
program p
complex, allocatable :: z(:)
allocate(z(2))
end program p
