! vybe-test: fortran/allocation_source/allocation_source_06
! origin: languages/fortran/tests/fortran/test_allocation_source.rs
program p
type t
 integer :: x
end type t
type(t), allocatable :: v
allocate(v, source=t(1))
end program p
