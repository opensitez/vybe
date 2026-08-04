! vybe-test: fortran/subroutine_extended/pure_square_of_six
! origin: languages/fortran/tests/fortran/test_subroutine_extended.rs
program t
if ((psquare(6)) /= 36) then
    print *, "FAIL: want [36] got [", psquare(6), "]"
    stop 1
end if
contains
pure function psquare(x) result(r)
integer, intent(in) :: x
integer :: r
r = x * x
end function psquare
end program t
