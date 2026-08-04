! vybe-test: fortran/oop/oop_property_like_13
! origin: languages/fortran/tests/fortran/test_oop.rs
type::t
integer::x
contains
procedure::getx
end type t
contains
integer function getx(this)
class(t)::this
getx=this%x
end function getx
