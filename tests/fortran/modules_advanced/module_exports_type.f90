! vybe-test: fortran/modules_advanced/module_exports_type
! origin: languages/fortran/tests/fortran/test_modules_advanced.rs

module geometry
    implicit none
    type :: Vector2D
        real :: x, y
    end type Vector2D
contains
    function length(v) result(r)
        type(Vector2D), intent(in) :: v
        real :: r
        r = sqrt(v%x**2 + v%y**2)
    end function length
end module geometry

program test
    use geometry
    type(Vector2D) :: v
    v%x = 3.0
    v%y = 4.0
    if ((length(v)) /= 5) then
    print *, "FAIL: want [5] got [", length(v), "]"
    stop 1
end if
end program test
