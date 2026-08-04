! vybe-test: fortran/allocate_statement/allocate_statement_05
! origin: languages/fortran/tests/fortran/test_allocate_statement.rs
program p
integer, allocatable :: x
allocate(x)
end program p
