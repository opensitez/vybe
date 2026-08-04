! vybe-test: fortran/oop/oop_encap_14
! origin: languages/fortran/tests/fortran/test_oop.rs
module m
type::t
private
integer::x
contains
procedure::setx
end type t
contains
subroutine setx(this,v)
class(t)::this
integer::v
this%x=v
end
end module m
