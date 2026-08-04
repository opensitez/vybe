! vybe-test: fortran/allocate_statement/allocate_statement_07
! origin: languages/fortran/tests/fortran/test_allocate_statement.rs
program p
logical, allocatable :: l(:)
allocate(l(2))
end program p
