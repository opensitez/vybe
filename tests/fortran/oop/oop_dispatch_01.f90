! vybe-test: fortran/oop/oop_dispatch_01
! origin: languages/fortran/tests/fortran/test_oop.rs
module m
type::b
contains
procedure::show
end type b
contains
subroutine show(this)
class(b)::this
end subroutine show
end module m
