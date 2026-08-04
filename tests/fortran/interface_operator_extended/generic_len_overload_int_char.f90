! vybe-test: fortran/interface_operator_extended/generic_len_overload_int_char
! origin: languages/fortran/tests/fortran/test_interface_operator_extended.rs
module glen
implicit none
interface span
module procedure span_int, span_char
end interface
contains
function span_int(n) result(r)
integer, intent(in) :: n
integer :: r
r = abs(n) + 1
end function span_int
function span_char(s) result(r)
character(len=*), intent(in) :: s
integer :: r
r = len_trim(s)
end function span_char
end module glen
program t
use glen
if ((span(-4)) /= 5) then
    print *, "FAIL: want [5] got [", span(-4), "]"
    stop 1
end if
if ((span('abcd')) /= 4) then
    print *, "FAIL: want [4] got [", span('abcd'), "]"
    stop 1
end if
end program t
