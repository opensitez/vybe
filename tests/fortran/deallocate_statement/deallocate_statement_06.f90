! vybe-test: fortran/deallocate_statement/deallocate_statement_06
! origin: languages/fortran/tests/fortran/test_deallocate_statement.rs
program driver
integer, pointer :: p(:)
allocate(p(3))
deallocate(p)
end program driver