! vybe-test: fortran/interface_operator_extended/generic_identity_overloads
! origin: languages/fortran/tests/fortran/test_interface_operator_extended.rs
module gid
implicit none
interface ident
module procedure ident_int, ident_char
end interface
contains
function ident_int(v) result(r)
integer, intent(in) :: v
integer :: r
r = v
end function ident_int
function ident_char(v) result(r)
character(len=*), intent(in) :: v
character(len=10) :: r
r = v
end function ident_char
end module gid
program t
use gid
if ((ident(42)) /= 42) then
    print *, "FAIL: want [42] got [", ident(42), "]"
    stop 1
end if
if (trim(trim(ident('z'))) /= "z") then
    print *, "FAIL: want [z] got [", trim(ident('z')), "]"
    stop 1
end if
end program t
