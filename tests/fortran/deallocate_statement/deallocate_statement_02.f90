! vybe-test: fortran/deallocate_statement/deallocate_statement_02
! origin: languages/fortran/tests/fortran/test_deallocate_statement.rs
program p
real, allocatable :: a(:,:)
allocate(a(2,2))
deallocate(a)
end program p
