! vybe-test: fortran/interface_operator_extended/generic_compare_strings_and_ints
! origin: languages/fortran/tests/fortran/test_interface_operator_extended.rs
module gcmp
implicit none
interface same
module procedure same_int, same_char
end interface
contains
function same_int(a, b) result(r)
integer, intent(in) :: a, b
logical :: r
r = a == b
end function same_int
function same_char(a, b) result(r)
character(len=*), intent(in) :: a, b
logical :: r
r = a == b
end function same_char
end module gcmp
program t
use gcmp
if ((same(3, 3)) .neqv. .true.) then
    print *, "FAIL: want [true] got [", same(3, 3), "]"
    stop 1
end if
if ((same('x', 'x')) .neqv. .true.) then
    print *, "FAIL: want [true] got [", same('x', 'x'), "]"
    stop 1
end if
end program t
