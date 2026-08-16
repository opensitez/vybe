! vybe-test: fortran/allocation/alloc_comp_08
! origin: languages/fortran/tests/fortran/test_allocation.rs
program p
type :: t
integer, allocatable :: a(:)
end type t
type(t)::x
allocate(x%a(3))
end program p
