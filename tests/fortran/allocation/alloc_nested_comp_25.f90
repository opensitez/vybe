! vybe-test: fortran/allocation/alloc_nested_comp_25
! origin: languages/fortran/tests/fortran/test_allocation.rs
program p
type :: t
integer, allocatable :: a(:)
end type t
type(t), allocatable :: x
allocate(x)
allocate(x%a(2))
end program p
