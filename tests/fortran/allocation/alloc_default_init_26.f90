! vybe-test: fortran/allocation/alloc_default_init_26
! origin: languages/fortran/tests/fortran/test_allocation.rs
type :: t
integer :: x=1
end type t
program p
type(t), allocatable :: a(:)
allocate(a(2))
end program p
