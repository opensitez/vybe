! vybe-test: fortran/subroutine_extended/elemental_pure_abs_sum
! origin: languages/fortran/tests/fortran/test_subroutine_extended.rs
program t
integer :: a(4), b(4)
a = [-1, 2, -3, 4]
b = eabsval(a)
if ((sum(b)) /= 10) then
    print *, "FAIL: want [10] got [", sum(b), "]"
    stop 1
end if
contains
elemental function eabsval(x) result(r)
integer, intent(in) :: x
integer :: r
if (x < 0) then
r = -x
else
r = x
end if
end function eabsval
end program t
