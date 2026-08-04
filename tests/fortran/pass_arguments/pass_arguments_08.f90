! vybe-test: fortran/pass_arguments/pass_arguments_08
! origin: languages/fortran/tests/fortran/test_pass_arguments.rs
module m
type::t
contains
procedure,pass::a
procedure,pass::b
end type
contains
subroutine a(this)
class(t)::this
end
subroutine b(this)
class(t)::this
end
end module m
