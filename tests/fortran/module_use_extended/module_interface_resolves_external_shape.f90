! vybe-test: fortran/module_use_extended/module_interface_resolves_external_shape
! origin: languages/fortran/tests/fortran/test_module_use_extended.rs
module iface_use
implicit none
interface
function extern_double(n) result(r)
integer, intent(in) :: n
integer :: r
end function extern_double
end interface
contains
function call_double(n) result(r)
integer, intent(in) :: n
integer :: r
r = extern_double(n)
end function call_double
end module iface_use
function extern_double(n) result(r)
integer, intent(in) :: n
integer :: r
r = n * 2
end function extern_double
program t
use iface_use
if ((call_double(6)) /= 12) then
    print *, "FAIL: want [12] got [", call_double(6), "]"
    stop 1
end if
end program t
