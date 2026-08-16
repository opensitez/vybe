! vybe-test: fortran/interfaces/if_generic_result_type_46
! origin: languages/fortran/tests/fortran/test_interfaces.rs
module m
interface g
module procedure gi, gr
end interface
contains
integer function gi(x)
integer, intent(in) :: x
gi = x
end function gi
real function gr(x)
real, intent(in) :: x
gr = x
end function gr
end module m
program t
use m
if (g(3) /= 3) then
    print *, "FAIL: want [3] got [", g(3), "]"
    stop 1
end if
if (abs(g(2.5) - 2.5) > 1.0e-6) then
    print *, "FAIL: want [2.5] got [", g(2.5), "]"
    stop 1
end if
end program t
