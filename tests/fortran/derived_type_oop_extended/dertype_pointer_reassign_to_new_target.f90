! vybe-test: fortran/derived_type_oop_extended/dertype_pointer_reassign_to_new_target
! origin: languages/fortran/tests/fortran/test_derived_type_oop_extended.rs
program t
type :: Pair
integer :: left = 0
integer, pointer :: right => null()
end type Pair
type(Pair) :: p
integer, target :: x = 3, y = 8
p%left = 1
p%right => x
if ((p%right) /= 3) then
    print *, "FAIL: want [3] got [", p%right, "]"
    stop 1
end if
p%right => y
if ((p%right) /= 8) then
    print *, "FAIL: want [8] got [", p%right, "]"
    stop 1
end if
end program t
