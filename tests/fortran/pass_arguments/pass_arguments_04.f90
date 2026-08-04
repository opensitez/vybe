! vybe-test: fortran/pass_arguments/pass_arguments_04
! origin: languages/fortran/tests/fortran/test_pass_arguments.rs
module m
type::t
contains
procedure,pass::s1
procedure,pass::s2
end type
contains
subroutine s1(this)
class(t)::this
end
subroutine s2(this)
class(t)::this
end
end module m
