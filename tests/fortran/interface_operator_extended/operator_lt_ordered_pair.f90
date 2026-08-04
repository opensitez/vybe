! vybe-test: fortran/interface_operator_extended/operator_lt_ordered_pair
! origin: languages/fortran/tests/fortran/test_interface_operator_extended.rs
module glt
implicit none
type :: Pair
integer :: a, b
end type Pair
interface operator(<)
module procedure lt_pair
end interface
contains
function lt_pair(x, y) result(r)
type(Pair), intent(in) :: x, y
logical :: r
r = x%a < y%a
end function lt_pair
end module glt
program t
use glt
type(Pair) :: p, q
p%a = 2
q%a = 5
if ((p < q) .neqv. .true.) then
    print *, "FAIL: want [true] got [", p < q, "]"
    stop 1
end if
end program t
