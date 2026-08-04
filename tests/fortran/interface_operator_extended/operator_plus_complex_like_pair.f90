! vybe-test: fortran/interface_operator_extended/operator_plus_complex_like_pair
! origin: languages/fortran/tests/fortran/test_interface_operator_extended.rs
module gcplx
implicit none
type :: Cplx
real :: re, im
end type Cplx
interface operator(+)
module procedure add_cplx
end interface
contains
function add_cplx(a, b) result(c)
type(Cplx), intent(in) :: a, b
type(Cplx) :: c
c%re = a%re + b%re
c%im = a%im + b%im
end function add_cplx
end module gcplx
program t
use gcplx
type(Cplx) :: a, b, c
a%re = 1.0; a%im = 2.0
b%re = 3.0; b%im = -1.0
c = a + b
if ((int(c%re)) /= 4) then
    print *, "FAIL: want [4] got [", int(c%re), "]"
    stop 1
end if
if ((int(c%im)) /= 1) then
    print *, "FAIL: want [1] got [", int(c%im), "]"
    stop 1
end if
end program t
