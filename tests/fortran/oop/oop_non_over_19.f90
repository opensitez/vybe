! vybe-test: fortran/oop/oop_non_over_19
! origin: languages/fortran/tests/fortran/test_oop.rs
module m
type::t
contains
procedure,non_overridable::s
end type t
contains
subroutine s(this)
class(t)::this
end subroutine s
end module m
