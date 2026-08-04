! vybe-test: fortran/allocation/alloc_class_17
! origin: languages/fortran/tests/fortran/test_allocation.rs
type :: t
integer::x
end type t
program p
class(t), allocatable :: obj
allocate(obj)
end program p
