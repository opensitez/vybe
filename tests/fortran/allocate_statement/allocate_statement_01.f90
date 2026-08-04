! vybe-test: fortran/allocate_statement/allocate_statement_01
! origin: languages/fortran/tests/fortran/test_allocate_statement.rs
program p
integer, allocatable :: a(:)
allocate(a(3))
end program p
