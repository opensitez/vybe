! vybe-test: fortran/subroutine_extended/elemental_scale_by_three_sum
! origin: languages/fortran/tests/fortran/test_subroutine_extended.rs
program t
integer :: a(3), b(3)
a = [1, 2, 3]
b = escale3(a)
if ((sum(b)) /= 18) then
    print *, "FAIL: want [18] got [", sum(b), "]"
    stop 1
end if
contains
elemental function escale3(x) result(r)
integer, intent(in) :: x
integer :: r
r = x * 3
end function escale3
end program t
