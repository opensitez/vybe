! vybe-test: fortran/module_use_extended/rename_combined_with_only_clause
! origin: languages/fortran/tests/fortran/test_module_use_extended.rs
module pack
implicit none
integer :: hidden = 1
integer :: visible = 2
contains
function pick(x) result(r)
integer, intent(in) :: x
integer :: r
r = x + hidden
end function pick
end module pack
program t
use pack, only: choose => pick
if ((choose(5)) /= 6) then
    print *, "FAIL: want [6] got [", choose(5), "]"
    stop 1
end if
end program t
