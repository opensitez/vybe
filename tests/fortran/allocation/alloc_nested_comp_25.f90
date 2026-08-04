! vybe-test: fortran/allocation/alloc_nested_comp_25
! origin: languages/fortran/tests/fortran/test_allocation.rs
type :: t
integer, allocatable :: a(:)
end type t
program p
type(t), allocatable :: x
allocate(x)
allocate(x%a(2))
end program p
