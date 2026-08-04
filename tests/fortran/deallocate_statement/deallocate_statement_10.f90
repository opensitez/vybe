! vybe-test: fortran/deallocate_statement/deallocate_statement_10
! origin: languages/fortran/tests/fortran/test_deallocate_statement.rs
program p
complex, allocatable :: z(:)
allocate(z(2))
deallocate(z)
end program p
