! vybe-test: fortran/nopass_arguments/nopass_arguments_05
! origin: languages/fortran/tests/fortran/test_nopass_arguments.rs
module m
type::t
contains
procedure,nopass::mk
end type
contains
function mk() result(r)
integer :: r
r=1
end
end module m
