! vybe-test: fortran/oop/oop_factory_04
! origin: languages/fortran/tests/fortran/test_oop.rs
module m
type::t
integer::x
end type t
contains
function make() result(r)
type(t)::r
r%x=1
end function make
end module m
