! vybe-test: fortran/module_use_extended/module_generic_len_overloads
! origin: languages/fortran/tests/fortran/test_module_use_extended.rs
module lenmod
implicit none
interface my_len
module procedure len_int, len_char
end interface
contains
function len_int(n) result(r)
integer, intent(in) :: n
integer :: r
r = 1
end function len_int
function len_char(s) result(r)
character(len=*), intent(in) :: s
integer :: r
r = len_trim(s)
end function len_char
end module lenmod
program t
use lenmod
if ((my_len(0)) /= 1) then
    print *, "FAIL: want [1] got [", my_len(0), "]"
    stop 1
end if
if ((my_len('abcd')) /= 4) then
    print *, "FAIL: want [4] got [", my_len('abcd'), "]"
    stop 1
end if
end program t
