! vybe-test: fortran/derived_types/dt_bound_op_04
! origin: languages/fortran/tests/fortran/test_derived_types.rs
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
