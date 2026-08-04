! vybe-test: fortran/pass_arguments/pass_arguments_06
! origin: languages/fortran/tests/fortran/test_pass_arguments.rs
module m
type::t
contains
procedure,pass::get
end type
contains
integer function get(this)
class(t)::this
get=1
end
end module m
