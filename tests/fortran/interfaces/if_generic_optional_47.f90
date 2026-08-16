! vybe-test: fortran/interfaces/if_generic_optional_47
! origin: languages/fortran/tests/fortran/test_interfaces.rs
module m
interface g
module procedure g_int, g_real
end interface
contains
integer function g_int(x, flag)
integer, intent(in) :: x
integer, intent(in), optional :: flag
g_int = x
if (present(flag)) g_int = x + flag
end function g_int
real function g_real(x, flag)
real, intent(in) :: x
logical, intent(in), optional :: flag
g_real = x
if (present(flag)) g_real = -x
end function g_real
end module m
program t
use m
if (g(3) /= 3) then
    print *, "FAIL: want [3] got [", g(3), "]"
    stop 1
end if
if (g(3, 4) /= 7) then
    print *, "FAIL: want [7] got [", g(3, 4), "]"
    stop 1
end if
if (abs(g(2.5) - 2.5) > 1.0e-6) then
    print *, "FAIL: want [2.5] got [", g(2.5), "]"
    stop 1
end if
if (abs(g(2.5, .true.) + 2.5) > 1.0e-6) then
    print *, "FAIL: want [-2.5] got [", g(2.5, .true.), "]"
    stop 1
end if
end program t
