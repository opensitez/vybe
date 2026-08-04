! vybe-test: fortran/allocation_source/allocation_source_06
! origin: languages/fortran/tests/fortran/test_allocation_source.rs
type t
 integer :: x
end type t
program p
type(t), allocatable :: v
allocate(v, source=t(1))
end program p
