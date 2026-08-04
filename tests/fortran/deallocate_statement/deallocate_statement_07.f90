! vybe-test: fortran/deallocate_statement/deallocate_statement_07
! origin: languages/fortran/tests/fortran/test_deallocate_statement.rs
program p
class(*), allocatable :: x
allocate(integer :: x)
deallocate(x)
end program p
