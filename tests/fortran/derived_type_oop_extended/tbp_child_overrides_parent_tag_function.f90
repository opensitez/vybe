! vybe-test: fortran/derived_type_oop_extended/tbp_child_overrides_parent_tag_function
! origin: languages/fortran/tests/fortran/test_derived_type_oop_extended.rs
program t
type :: Base
contains
procedure :: tag
end type Base
type, extends(Base) :: Child
contains
procedure :: tag => child_tag
end type Child
type(Child) :: c
if ((c%tag()) /= 2) then
    print *, "FAIL: want [2] got [", c%tag(), "]"
    stop 1
end if
contains
integer function tag(self) result(v)
class(Base), intent(in) :: self
v = 1
end function tag
integer function child_tag(self) result(v)
class(Child), intent(in) :: self
v = 2
end function child_tag
end program t
