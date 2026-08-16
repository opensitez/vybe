! vybe-test: fortran/allocation_mold/allocation_mold_07
! origin: languages/fortran/tests/fortran/test_allocation_mold.rs
program p
type t
 integer :: x
end type t
type(t), allocatable :: a,b
allocate(b)
allocate(a, mold=b)
end program p
