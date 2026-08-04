! vybe-test: fortran/interfaces/if_nopass_14
! origin: languages/fortran/tests/fortran/test_interfaces.rs
module m
type::t
contains
procedure,nopass::s
end type
contains
subroutine s()
end
end module m
