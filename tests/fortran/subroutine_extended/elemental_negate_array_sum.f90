! vybe-test: fortran/subroutine_extended/elemental_negate_array_sum
! origin: languages/fortran/tests/fortran/test_subroutine_extended.rs
program t
integer :: a(3), b(3)
a = [1, 2, 3]
b = eneg(a)
if ((sum(b)) /= -6) then
    print *, "FAIL: want [-6] got [", sum(b), "]"
    stop 1
end if
contains
elemental function eneg(x) result(r)
integer, intent(in) :: x
integer :: r
r = -x
end function eneg
end program t
