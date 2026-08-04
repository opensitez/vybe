! vybe-test: fortran/nopass_arguments/nopass_arguments_09
! origin: languages/fortran/tests/fortran/test_nopass_arguments.rs
module m
type::t
contains
procedure,nopass::s
end type
contains
subroutine s(c)
character(len=*)::c
end
end module m
