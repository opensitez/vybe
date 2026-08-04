! vybe-test: fortran/nopass_arguments/nopass_arguments_06
! origin: languages/fortran/tests/fortran/test_nopass_arguments.rs
module m
type::t
contains
procedure,nopass::show
end type
contains
subroutine show()
print *,1
end
end module m
