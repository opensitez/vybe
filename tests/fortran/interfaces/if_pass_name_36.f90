! vybe-test: fortran/interfaces/if_pass_name_36
! origin: languages/fortran/tests/fortran/test_interfaces.rs
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
