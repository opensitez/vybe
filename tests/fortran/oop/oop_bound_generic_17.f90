! vybe-test: fortran/oop/oop_bound_generic_17
! origin: languages/fortran/tests/fortran/test_oop.rs
module m
type::t
contains
generic::g=>s
procedure::s
end type t
contains
subroutine s(this)
class(t)::this
end subroutine s
end module m
