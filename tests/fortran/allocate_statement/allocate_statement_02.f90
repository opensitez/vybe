! vybe-test: fortran/allocate_statement/allocate_statement_02
! origin: languages/fortran/tests/fortran/test_allocate_statement.rs
program p
real, allocatable :: a(:,:)
allocate(a(2,2))
end program p
