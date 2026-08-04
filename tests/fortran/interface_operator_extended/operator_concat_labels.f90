! vybe-test: fortran/interface_operator_extended/operator_concat_labels
! origin: languages/fortran/tests/fortran/test_interface_operator_extended.rs
module glabel
implicit none
type :: Label
character(len=8) :: text
end type Label
interface operator(//)
module procedure concat_label
end interface
contains
function concat_label(a, b) result(c)
type(Label), intent(in) :: a, b
type(Label) :: c
c%text = trim(a%text) // trim(b%text)
end function concat_label
end module glabel
program t
use glabel
type(Label) :: a, b, c
a%text = 'foo'
b%text = 'bar'
c = a // b
if (trim(trim(c%text)) /= "foobar") then
    print *, "FAIL: want [foobar] got [", trim(c%text), "]"
    stop 1
end if
end program t
