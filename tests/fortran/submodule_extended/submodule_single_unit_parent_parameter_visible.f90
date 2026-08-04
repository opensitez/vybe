! vybe-test: fortran/submodule_extended/submodule_single_unit_parent_parameter_visible
! origin: languages/fortran/tests/fortran/test_submodule_extended.rs
module host
implicit none
integer, parameter :: TAG = 99
interface
module function bump(x) result(r)
integer, intent(in) :: x
integer :: r
end function bump
end interface
end module host
submodule (host) host_impl
contains
module function bump(x) result(r)
integer, intent(in) :: x
integer :: r
r = x + TAG
end function bump
end submodule host_impl
program t
use host
if ((TAG) /= 99) then
    print *, "FAIL: want [99] got [", TAG, "]"
    stop 1
end if
end program t
