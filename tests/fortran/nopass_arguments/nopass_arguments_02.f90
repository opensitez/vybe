! vybe-test: fortran/nopass_arguments/nopass_arguments_02
! origin: languages/fortran/tests/fortran/test_nopass_arguments.rs
module m
type::t
contains
procedure,nopass::f
end type
contains
integer function f()
f=1
end
end module m
