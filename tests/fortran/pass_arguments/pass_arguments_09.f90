! vybe-test: fortran/pass_arguments/pass_arguments_09
! origin: languages/fortran/tests/fortran/test_pass_arguments.rs
module m
type::t
contains
procedure,pass::s
end type
contains
subroutine s(this)
class(t), intent(inout) :: this
end
end module m
