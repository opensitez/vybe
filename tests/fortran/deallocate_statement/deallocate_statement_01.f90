! vybe-test: fortran/deallocate_statement/deallocate_statement_01
! origin: languages/fortran/tests/fortran/test_deallocate_statement.rs
program p
integer, allocatable :: a(:)
allocate(a(3))
deallocate(a)
end program p
