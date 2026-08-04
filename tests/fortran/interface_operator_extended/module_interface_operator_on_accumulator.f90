! vybe-test: fortran/interface_operator_extended/module_interface_operator_on_accumulator
! origin: languages/fortran/tests/fortran/test_interface_operator_extended.rs
module gacc2
implicit none
type :: Acc
integer :: total
end type Acc
interface operator(+)
module procedure add_acc
end interface
contains
function add_acc(a, b) result(c)
type(Acc), intent(in) :: a, b
type(Acc) :: c
c%total = a%total + b%total
end function add_acc
end module gacc2
program t
use gacc2
type(Acc) :: seed, step, out
seed%total = 10
step%total = 5
out = seed + step
if ((out%total) /= 15) then
    print *, "FAIL: want [15] got [", out%total, "]"
    stop 1
end if
end program t
