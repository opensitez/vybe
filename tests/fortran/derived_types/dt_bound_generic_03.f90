! vybe-test: fortran/derived_types/dt_bound_generic_03
! origin: languages/fortran/tests/fortran/test_derived_types.rs
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
