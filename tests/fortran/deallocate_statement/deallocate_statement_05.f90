! vybe-test: fortran/deallocate_statement/deallocate_statement_05
! origin: languages/fortran/tests/fortran/test_deallocate_statement.rs
program p
integer, allocatable :: a(:), b(:)
allocate(a(2),b(2))
deallocate(a,b)
end program p
