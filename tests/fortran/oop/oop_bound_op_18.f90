! vybe-test: fortran/oop/oop_bound_op_18
! origin: languages/fortran/tests/fortran/test_oop.rs
module m
type::t
contains
procedure::add
generic::operator(+)=>add
end type t
contains
integer function add(this,other)
class(t)::this
class(t)::other
add=0
end function add
end module m
