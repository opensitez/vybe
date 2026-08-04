! vybe-test: fortran/interface_operator_extended/generic_add_int_and_real
! origin: languages/fortran/tests/fortran/test_interface_operator_extended.rs
module gadd
implicit none
interface add_generic
module procedure add_int, add_real
end interface
contains
function add_int(a, b) result(r)
integer, intent(in) :: a, b
integer :: r
r = a + b
end function add_int
function add_real(a, b) result(r)
real, intent(in) :: a, b
real :: r
r = a + b
end function add_real
end module gadd
program t
use gadd
if ((add_generic(2, 3)) /= 5) then
    print *, "FAIL: want [5] got [", add_generic(2, 3), "]"
    stop 1
end if
if ((int(add_generic(1.5, 2.5))) /= 4) then
    print *, "FAIL: want [4] got [", int(add_generic(1.5, 2.5)), "]"
    stop 1
end if
end program t
