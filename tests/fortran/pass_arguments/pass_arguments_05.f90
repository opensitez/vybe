! vybe-test: fortran/pass_arguments/pass_arguments_05
! origin: languages/fortran/tests/fortran/test_pass_arguments.rs
module m
type::t
contains
procedure,pass::show
end type
contains
subroutine show(this)
class(t)::this
print *,1
end
end module m
