! vybe-test: fortran/pass_arguments/pass_arguments_03
! origin: languages/fortran/tests/fortran/test_pass_arguments.rs
module m
type::t
contains
procedure,pass(arg)::s
end type
contains
subroutine s(arg)
class(t)::arg
end
end module m
