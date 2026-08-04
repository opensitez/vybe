! vybe-test: fortran/subroutine_extended/pure_integer_max_pair
! origin: languages/fortran/tests/fortran/test_subroutine_extended.rs
program t
if ((pmax(12, 9)) /= 12) then
    print *, "FAIL: want [12] got [", pmax(12, 9), "]"
    stop 1
end if
contains
pure function pmax(a, b) result(m)
integer, intent(in) :: a, b
integer :: m
if (a >= b) then
m = a
else
m = b
end if
end function pmax
end program t
