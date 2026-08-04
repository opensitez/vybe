! vybe-test: fortran/pointer_alloc_extended/associated_derived_type_field_initially_false
! origin: languages/fortran/tests/fortran/test_pointer_alloc_extended.rs
program t
type :: Node
integer :: id
integer, pointer :: child => null()
end type Node
type(Node) :: n
n%id = 1
if ((associated(n%child)) .neqv. .false.) then
    print *, "FAIL: want [false] got [", associated(n%child), "]"
    stop 1
end if
end program t
