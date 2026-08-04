! vybe-test: fortran/derived_type_oop_extended/dertype_pointer_targets_scalar_member
! origin: languages/fortran/tests/fortran/test_derived_type_oop_extended.rs
program t
type :: Node
integer :: key = 0
integer, pointer :: link => null()
end type Node
type(Node), target :: a, b
a%key = 11
b%key = 22
a%link => b
if ((a%link%key) /= 22) then
    print *, "FAIL: want [22] got [", a%link%key, "]"
    stop 1
end if
end program t
