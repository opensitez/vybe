! vybe-test: fortran/submodules_advanced/submodule_uses_parent_type
! origin: languages/fortran/tests/fortran/test_submodules_advanced.rs

module geometry_iface
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
end module geometry_iface

submodule (geometry_iface) geometry_impl
    implicit none
contains
    module function distance(a, b) result(d)
        type(Point), intent(in) :: a, b
        real :: d
        d = sqrt((a%x - b%x)**2 + (a%y - b%y)**2)
    end function distance
end submodule geometry_impl

program test
    use geometry_iface
    type(Point) :: p1, p2
    p1 = Point(0.0, 0.0)
    p2 = Point(3.0, 4.0)
    if ((int(distance(p1, p2))) /= 5) then
    print *, "FAIL: want [5] got [", int(distance(p1, p2)), "]"
    stop 1
end if
end program test
