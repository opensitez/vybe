! vybe-test: fortran/oop/oop_self_ref_11
! origin: languages/fortran/tests/fortran/test_oop.rs
module m
type::t
contains
procedure::show
end type t
contains
subroutine show(this)
class(t)::this
print *,1
end subroutine show
end module m
