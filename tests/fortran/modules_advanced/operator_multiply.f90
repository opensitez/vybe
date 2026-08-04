! vybe-test: fortran/modules_advanced/operator_multiply
! origin: languages/fortran/tests/fortran/test_modules_advanced.rs

module vec_mod
    implicit none
    type :: Vec3
        real :: x, y, z
    end type Vec3
    interface operator(*)
        module procedure scale_vec
    end interface
contains
    function scale_vec(s, v) result(r)
        real, intent(in) :: s
        type(Vec3), intent(in) :: v
        type(Vec3) :: r
        r%x = s * v%x; r%y = s * v%y; r%z = s * v%z
    end function scale_vec
end module vec_mod

program test
    use vec_mod
    type(Vec3) :: v, r
    v = Vec3(1.0, 2.0, 3.0)
    r = 2.0 * v
    if ((r%x) /= 2) then
    print *, "FAIL: want [2] got [", r%x, "]"
    stop 1
end if
end program test
