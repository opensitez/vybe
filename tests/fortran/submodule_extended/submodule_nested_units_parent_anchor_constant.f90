! vybe-test: fortran/submodule_extended/submodule_nested_units_parent_anchor_constant
! origin: languages/fortran/tests/fortran/test_submodule_extended.rs
module anchor_iface
implicit none
integer, parameter :: ANCHOR = 77
interface
module function lift(x) result(r)
integer, intent(in) :: x
integer :: r
end function lift
end interface
end module anchor_iface
submodule (anchor_iface) mid_iface
interface
module function helper(x) result(r)
integer, intent(in) :: x
integer :: r
end function helper
end interface
end submodule mid_iface
submodule (anchor_iface:mid_iface) bot_impl
contains
module function lift(x) result(r)
integer, intent(in) :: x
integer :: r
r = x + ANCHOR
end function lift
module function helper(x) result(r)
integer, intent(in) :: x
integer :: r
r = x + 1
end function helper
end submodule bot_impl
program t
use anchor_iface
if ((ANCHOR) /= 77) then
    print *, "FAIL: want [77] got [", ANCHOR, "]"
    stop 1
end if
end program t
