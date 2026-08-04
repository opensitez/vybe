! vybe-test: fortran/submodule_extended/submodule_parent_type_ctor_without_calling_iface
! origin: languages/fortran/tests/fortran/test_submodule_extended.rs
module geom_iface
implicit none
type :: Point
real :: x, y
end type Point
interface
module function distance(a, b) result(d)
type(Point), intent(in) :: a, b
real :: d
end function distance
end interface
end module geom_iface
submodule (geom_iface) geom_impl
contains
module function distance(a, b) result(d)
type(Point), intent(in) :: a, b
real :: d
d = sqrt((a%x - b%x)**2 + (a%y - b%y)**2)
end function distance
end submodule geom_impl
program t
use geom_iface
type(Point) :: p
p%x = 3.0
p%y = 4.0
if ((int(p%x)) /= 3) then
    print *, "FAIL: want [3] got [", int(p%x), "]"
    stop 1
end if
end program t
