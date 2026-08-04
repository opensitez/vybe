! vybe-test: fortran/module_use_extended/rename_two_symbols_in_one_use
! origin: languages/fortran/tests/fortran/test_module_use_extended.rs
module pair
implicit none
integer :: first = 10
integer :: second = 3
contains
function sum_pair() result(r)
integer :: r
r = first + second
end function sum_pair
end module pair
program t
use pair, a => first, b => second, total => sum_pair
if ((total()) /= 13) then
    print *, "FAIL: want [13] got [", total(), "]"
    stop 1
end if
end program t
