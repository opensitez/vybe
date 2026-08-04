! vybe-test: fortran/fortran2003_extended/alloc_comp_module_derived_field
! origin: languages/fortran/tests/fortran/test_fortran2003_extended.rs
module nodes
implicit none
type :: Node
integer :: value = 0
end type Node
type :: List
type(Node), allocatable :: items(:)
end type List
end module nodes
program t
use nodes
type(List) :: lst
allocate(lst%items(1))
lst%items(1)%value = 42
if ((lst%items(1)%value) /= 42) then
    print *, "FAIL: want [42] got [", lst%items(1)%value, "]"
    stop 1
end if
end program t
