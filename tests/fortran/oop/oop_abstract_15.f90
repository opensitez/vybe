! vybe-test: fortran/oop/oop_abstract_15
! origin: languages/fortran/tests/fortran/test_oop.rs
type,abstract::t
contains
procedure(p),deferred::run
end type t
abstract interface
subroutine p(this)
import t
class(t)::this
end
end interface
