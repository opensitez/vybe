! vybe-test: fortran/interface_operator_extended/operator_minus_unary_on_metric
! origin: languages/fortran/tests/fortran/test_interface_operator_extended.rs
module gmet
implicit none
type :: Metric
integer :: mm
end type Metric
interface operator(-)
module procedure neg_metric
end interface
contains
function neg_metric(a) result(b)
type(Metric), intent(in) :: a
type(Metric) :: b
b%mm = -a%mm
end function neg_metric
end module gmet
program t
use gmet
type(Metric) :: m, n
m%mm = 15
n = -m
if ((n%mm) /= -15) then
    print *, "FAIL: want [-15] got [", n%mm, "]"
    stop 1
end if
end program t
