! vybe-test: fortran/attributes/attr_deferred_12
! origin: languages/fortran/tests/fortran/test_attributes.rs
type, abstract :: t
contains
procedure(p),deferred::s
end type t
abstract interface
subroutine p(this)
import t
class(t)::this
end
end interface
