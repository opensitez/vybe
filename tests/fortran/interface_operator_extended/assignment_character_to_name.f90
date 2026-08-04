! vybe-test: fortran/interface_operator_extended/assignment_character_to_name
! origin: languages/fortran/tests/fortran/test_interface_operator_extended.rs
module gname
implicit none
type :: Name
character(len=10) :: text
end type Name
interface assignment(=)
module procedure char_to_name
end interface
contains
subroutine char_to_name(dest, src)
type(Name), intent(out) :: dest
character(len=*), intent(in) :: src
dest%text = src
end subroutine char_to_name
end module gname
program t
use gname
type(Name) :: n
n = 'Fortran'
if (trim(trim(n%text)) /= "Fortran") then
    print *, "FAIL: want [Fortran] got [", trim(n%text), "]"
    stop 1
end if
end program t
