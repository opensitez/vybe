! vybe-test: fortran/submodule_extended/submodule_use_rename_binding_from_submodule
! origin: languages/fortran/tests/fortran/test_submodule_extended.rs
module scale_iface
implicit none
interface
module function scale(x) result(r)
integer, intent(in) :: x
integer :: r
end function scale
end interface
end module scale_iface
submodule (scale_iface) scale_impl
contains
module function scale(x) result(r)
integer, intent(in) :: x
integer :: r
r = x * 3
end function scale
end submodule scale_impl
program t
use scale_iface, only: triple => scale
if ((triple(4)) /= 12) then
    print *, "FAIL: want [12] got [", triple(4), "]"
    stop 1
end if
end program t
