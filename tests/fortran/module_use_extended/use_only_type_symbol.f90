! vybe-test: fortran/module_use_extended/use_only_type_symbol
! origin: languages/fortran/tests/fortran/test_module_use_extended.rs
module geom
implicit none
type :: Point
real :: x, y
end type Point
contains
function origin() result(p)
type(Point) :: p
p%x = 0.0
p%y = 0.0
end function origin
end module geom
program t
use geom, only: Point
type(Point) :: p
p%x = 3.0
p%y = 4.0
if ((int(p%x)) /= 3) then
    print *, "FAIL: want [3] got [", int(p%x), "]"
    stop 1
end if
end program t
