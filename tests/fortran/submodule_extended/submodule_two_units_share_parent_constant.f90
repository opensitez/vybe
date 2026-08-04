! vybe-test: fortran/submodule_extended/submodule_two_units_share_parent_constant
! origin: languages/fortran/tests/fortran/test_submodule_extended.rs
module pair_iface
implicit none
integer, parameter :: BASE = 3
interface
module function add1(x) result(r)
integer, intent(in) :: x
integer :: r
end function add1
module function add2(x) result(r)
integer, intent(in) :: x
integer :: r
end function add2
end interface
end module pair_iface
submodule (pair_iface) pair_a
contains
module function add1(x) result(r)
integer, intent(in) :: x
integer :: r
r = x + BASE
end function add1
end submodule pair_a
submodule (pair_iface) pair_b
contains
module function add2(x) result(r)
integer, intent(in) :: x
integer :: r
r = x + BASE + BASE
end function add2
end submodule pair_b
program t
use pair_iface
if ((BASE) /= 3) then
    print *, "FAIL: want [3] got [", BASE, "]"
    stop 1
end if
end program t
