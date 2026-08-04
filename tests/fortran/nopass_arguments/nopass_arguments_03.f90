! vybe-test: fortran/nopass_arguments/nopass_arguments_03
! origin: languages/fortran/tests/fortran/test_nopass_arguments.rs
module m
type::t
contains
procedure,nopass::s1
procedure,nopass::s2
end type
contains
subroutine s1()
end
subroutine s2()
end
end module m
