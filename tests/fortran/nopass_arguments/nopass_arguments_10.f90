! vybe-test: fortran/nopass_arguments/nopass_arguments_10
! origin: languages/fortran/tests/fortran/test_nopass_arguments.rs
module m
type::t
contains
procedure,nopass::s
end type
contains
subroutine s(l)
logical::l
end
end module m
