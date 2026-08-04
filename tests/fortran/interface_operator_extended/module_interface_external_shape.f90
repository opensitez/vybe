! vybe-test: fortran/interface_operator_extended/module_interface_external_shape
! origin: languages/fortran/tests/fortran/test_interface_operator_extended.rs
module gext
implicit none
interface
function extern_triple(n) result(r)
integer, intent(in) :: n
integer :: r
end function extern_triple
end interface
contains
function via_iface(n) result(r)
integer, intent(in) :: n
integer :: r
r = extern_triple(n)
end function via_iface
end module gext
function extern_triple(n) result(r)
integer, intent(in) :: n
integer :: r
r = n * 3
end function extern_triple
program t
use gext
if ((via_iface(4)) /= 12) then
    print *, "FAIL: want [12] got [", via_iface(4), "]"
    stop 1
end if
end program t
