! vybe-test: fortran/pass_arguments/pass_arguments_02
! origin: languages/fortran/tests/fortran/test_pass_arguments.rs
module m
type::t
contains
procedure,pass(self)::s
end type
contains
subroutine s(self)
class(t)::self
end
end module m
