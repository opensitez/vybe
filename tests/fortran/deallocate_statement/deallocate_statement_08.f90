! vybe-test: fortran/deallocate_statement/deallocate_statement_08
! origin: languages/fortran/tests/fortran/test_deallocate_statement.rs
program p
type t
 integer :: x
end type t
type(t), allocatable :: a(:)
allocate(a(2))
deallocate(a)
end program p
