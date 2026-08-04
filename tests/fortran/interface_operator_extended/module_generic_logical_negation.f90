! vybe-test: fortran/interface_operator_extended/module_generic_logical_negation
! origin: languages/fortran/tests/fortran/test_interface_operator_extended.rs
module gnot
implicit none
interface flip
module procedure flip_log, flip_int
end interface
contains
function flip_log(v) result(r)
logical, intent(in) :: v
logical :: r
r = .not. v
end function flip_log
function flip_int(v) result(r)
integer, intent(in) :: v
integer :: r
r = -v
end function flip_int
end module gnot
program t
use gnot
if ((flip(.true.)) .neqv. .false.) then
    print *, "FAIL: want [false] got [", flip(.true.), "]"
    stop 1
end if
if ((flip(8)) /= -8) then
    print *, "FAIL: want [-8] got [", flip(8), "]"
    stop 1
end if
end program t
