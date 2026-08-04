! vybe-test: fortran/subroutine_extended/pure_cube_of_three
! origin: languages/fortran/tests/fortran/test_subroutine_extended.rs
program t
if ((pcube(3)) /= 27) then
    print *, "FAIL: want [27] got [", pcube(3), "]"
    stop 1
end if
contains
pure function pcube(x) result(r)
integer, intent(in) :: x
integer :: r
r = x * x * x
end function pcube
end program t
