! vybe-test: fortran/allocation/alloc_default_init_26
! origin: languages/fortran/tests/fortran/test_allocation.rs
program p
type :: t
integer :: x=1
end type t
type(t), allocatable :: a(:)
allocate(a(2))
end program p
