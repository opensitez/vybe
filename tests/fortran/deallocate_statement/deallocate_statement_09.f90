! vybe-test: fortran/deallocate_statement/deallocate_statement_09
! origin: languages/fortran/tests/fortran/test_deallocate_statement.rs
program p
logical, allocatable :: l(:)
allocate(l(2))
deallocate(l)
end program p
