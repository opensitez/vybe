! vybe-test: fortran/pass_arguments/pass_arguments_07
! origin: languages/fortran/tests/fortran/test_pass_arguments.rs
module m
type::t
contains
procedure,pass::set
end type
contains
subroutine set(this,x)
class(t)::this
integer::x
end
end module m
