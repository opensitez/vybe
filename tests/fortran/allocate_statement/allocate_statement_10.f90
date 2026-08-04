! vybe-test: fortran/allocate_statement/allocate_statement_10
! origin: languages/fortran/tests/fortran/test_allocate_statement.rs
program p
class(*), allocatable :: x
allocate(integer :: x)
end program p
