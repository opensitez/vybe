! vybe-test: fortran/interface_operator_extended/operator_eq_on_character_wrapper
! origin: languages/fortran/tests/fortran/test_interface_operator_extended.rs
module gstr
implicit none
type :: Str
character(len=6) :: data
end type Str
interface operator(==)
module procedure eq_str
end interface
contains
function eq_str(a, b) result(r)
type(Str), intent(in) :: a, b
logical :: r
r = trim(a%data) == trim(b%data)
end function eq_str
end module gstr
program t
use gstr
type(Str) :: a, b
a%data = 'hi'
b%data = 'hi'
if ((a == b) .neqv. .true.) then
    print *, "FAIL: want [true] got [", a == b, "]"
    stop 1
end if
end program t
