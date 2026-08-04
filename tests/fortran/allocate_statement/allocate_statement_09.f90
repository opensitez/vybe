! vybe-test: fortran/allocate_statement/allocate_statement_09
! origin: languages/fortran/tests/fortran/test_allocate_statement.rs
program p
type t
 integer :: x
end type t
type(t), allocatable :: a(:)
allocate(a(2))
end program p
