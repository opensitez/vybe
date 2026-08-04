! vybe-test: fortran/subroutine_extended/elemental_add_ten_second_element
! origin: languages/fortran/tests/fortran/test_subroutine_extended.rs
program t
integer :: a(3), b(3)
a = [1, 2, 3]
b = eplus10(a)
if ((b(2)) /= 12) then
    print *, "FAIL: want [12] got [", b(2), "]"
    stop 1
end if
contains
elemental function eplus10(x) result(r)
integer, intent(in) :: x
integer :: r
r = x + 10
end function eplus10
end program t
