! vybe-test: fortran/submodule_extended/submodule_parent_contained_helper_callable_in_submodule
! origin: languages/fortran/tests/fortran/test_submodule_extended.rs
module helper_iface
implicit none
interface
module function boosted(x) result(r)
integer, intent(in) :: x
integer :: r
end function boosted
end interface
contains
integer function local_offset()
local_offset = 2
end function local_offset
end module helper_iface
submodule (helper_iface) helper_impl
contains
module function boosted(x) result(r)
integer, intent(in) :: x
integer :: r
r = x + local_offset()
end function boosted
end submodule helper_impl
program t
use helper_iface
if ((boosted(5)) /= 7) then
    print *, "FAIL: want [7] got [", boosted(5), "]"
    stop 1
end if
end program t
