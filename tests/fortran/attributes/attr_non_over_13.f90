! vybe-test: fortran/attributes/attr_non_over_13
! origin: languages/fortran/tests/fortran/test_attributes.rs
module m
type::t
contains
procedure,non_overridable::s
end type
contains
subroutine s(this)
class(t)::this
end
end module m
