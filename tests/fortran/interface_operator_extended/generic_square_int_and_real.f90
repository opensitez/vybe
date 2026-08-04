! vybe-test: fortran/interface_operator_extended/generic_square_int_and_real
! origin: languages/fortran/tests/fortran/test_interface_operator_extended.rs
module gsqr
implicit none
interface square
module procedure square_int, square_real
end interface
contains
function square_int(x) result(r)
integer, intent(in) :: x
integer :: r
r = x * x
end function square_int
function square_real(x) result(r)
real, intent(in) :: x
real :: r
r = x * x
end function square_real
end module gsqr
program t
use gsqr
if ((square(6)) /= 36) then
    print *, "FAIL: want [36] got [", square(6), "]"
    stop 1
end if
if ((int(square(2.5))) /= 6) then
    print *, "FAIL: want [6] got [", int(square(2.5)), "]"
    stop 1
end if
end program t
