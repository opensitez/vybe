! vybe-test: fortran/attributes/attr_pointer_comp_25
! origin: languages/fortran/tests/fortran/test_attributes.rs
program driver
type :: t
integer, pointer :: p
end type t
type(t) :: obj
integer, target :: v
v = 9
obj%p => v
obj%p = 11
if (v /= 11) then
    print *, "FAIL: want [11] got [", v, "]"
    stop 1
end if
end program driver
