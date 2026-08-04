! vybe-test: fortran/allocation_semantics/as_13
! origin: languages/fortran/tests/fortran/test_allocation_semantics.rs
type :: t
integer, allocatable :: a(:)
end type t
program p
type(t) :: x
allocate(x%a(2))
end program p
