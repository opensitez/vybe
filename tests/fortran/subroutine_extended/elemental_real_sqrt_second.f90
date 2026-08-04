! vybe-test: fortran/subroutine_extended/elemental_real_sqrt_second
! origin: languages/fortran/tests/fortran/test_subroutine_extended.rs
program t
real :: a(3), b(3)
a = [4.0, 9.0, 16.0]
b = esqrt(a)
if ((b(2)) /= 3) then
    print *, "FAIL: want [3] got [", b(2), "]"
    stop 1
end if
contains
elemental function esqrt(x) result(r)
real, intent(in) :: x
real :: r
r = sqrt(x)
end function esqrt
end program t
