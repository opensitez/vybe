! vybe-test: fortran/interface_operator_extended/assignment_real_to_metric
! origin: languages/fortran/tests/fortran/test_interface_operator_extended.rs
module gmetric
implicit none
type :: Metric
real :: value
end type Metric
interface assignment(=)
module procedure real_to_metric
end interface
contains
subroutine real_to_metric(dest, src)
type(Metric), intent(out) :: dest
real, intent(in) :: src
dest%value = src * 2.0
end subroutine real_to_metric
end module gmetric
program t
use gmetric
type(Metric) :: m
m = 3.0
if ((int(m%value)) /= 6) then
    print *, "FAIL: want [6] got [", int(m%value), "]"
    stop 1
end if
end program t
