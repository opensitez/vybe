! vybe-test: fortran/interfaces/if_pass_13
! origin: languages/fortran/tests/fortran/test_interfaces.rs
module m
type::t
contains
procedure,pass::s
end type
contains
subroutine s(this)
class(t)::this
end
end module m
