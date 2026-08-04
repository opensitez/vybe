! vybe-test: fortran/subroutine_extended/elemental_double_array_sum
! origin: languages/fortran/tests/fortran/test_subroutine_extended.rs
program t
integer :: a(4), b(4)
a = [1, 2, 3, 4]
b = edouble(a)
if ((sum(b)) /= 20) then
    print *, "FAIL: want [20] got [", sum(b), "]"
    stop 1
end if
contains
elemental function edouble(x) result(r)
integer, intent(in) :: x
integer :: r
r = x * 2
end function edouble
end program t
