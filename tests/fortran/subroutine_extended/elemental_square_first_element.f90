! vybe-test: fortran/subroutine_extended/elemental_square_first_element
! origin: languages/fortran/tests/fortran/test_subroutine_extended.rs
program t
integer :: a(3), b(3)
a = [2, 3, 5]
b = esquare(a)
if ((b(1)) /= 4) then
    print *, "FAIL: want [4] got [", b(1), "]"
    stop 1
end if
contains
elemental function esquare(x) result(r)
integer, intent(in) :: x
integer :: r
r = x * x
end function esquare
end program t
