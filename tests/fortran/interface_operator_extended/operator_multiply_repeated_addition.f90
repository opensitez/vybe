! vybe-test: fortran/interface_operator_extended/operator_multiply_repeated_addition
! origin: languages/fortran/tests/fortran/test_interface_operator_extended.rs
module gscale
implicit none
type :: Weight
integer :: grams
end type Weight
interface operator(*)
module procedure scale_weight
end interface
contains
function scale_weight(w, n) result(r)
type(Weight), intent(in) :: w
integer, intent(in) :: n
type(Weight) :: r
r%grams = w%grams * n
end function scale_weight
end module gscale
program t
use gscale
type(Weight) :: w, r
w%grams = 5
r = w * 4
if ((r%grams) /= 20) then
    print *, "FAIL: want [20] got [", r%grams, "]"
    stop 1
end if
end program t
