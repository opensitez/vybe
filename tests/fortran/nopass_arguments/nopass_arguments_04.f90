! vybe-test: fortran/nopass_arguments/nopass_arguments_04
! origin: languages/fortran/tests/fortran/test_nopass_arguments.rs
module m
type::t
contains
procedure,nopass::set
end type
contains
subroutine set(x)
integer::x
end
end module m
