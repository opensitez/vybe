! vybe-test: fortran/interface_operator_extended/operator_plus_vector1d_lengths
! origin: languages/fortran/tests/fortran/test_interface_operator_extended.rs
module gvec
implicit none
type :: Vec
integer :: n
end type Vec
interface operator(+)
module procedure add_vec
end interface
contains
function add_vec(a, b) result(c)
type(Vec), intent(in) :: a, b
type(Vec) :: c
c%n = a%n + b%n
end function add_vec
end module gvec
program t
use gvec
type(Vec) :: u, v, w
u%n = 5
v%n = 7
w = u + v
if ((w%n) /= 12) then
    print *, "FAIL: want [12] got [", w%n, "]"
    stop 1
end if
end program t
