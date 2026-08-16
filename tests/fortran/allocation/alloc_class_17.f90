! vybe-test: fortran/allocation/alloc_class_17
! origin: languages/fortran/tests/fortran/test_allocation.rs
program p
type :: t
integer::x
end type t
class(t), allocatable :: obj
allocate(obj)
end program p
