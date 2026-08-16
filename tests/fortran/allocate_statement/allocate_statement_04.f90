! vybe-test: fortran/allocate_statement/allocate_statement_04
! origin: languages/fortran/tests/fortran/test_allocate_statement.rs
program driver
integer, pointer :: p(:)
allocate(p(3))
end program driver