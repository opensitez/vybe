! vybe-test: fortran/control/ctrl_return_09
! origin: languages/fortran/tests/fortran/test_control.rs
program test
integer :: x
call s(x)
if ((x) /= 1) then
    print *, "FAIL: want [1] got [", x, "]"
    stop 1
end if
contains
subroutine s(a)
    integer, intent(out) :: a
    a = 1
    return
    a = 2
end subroutine s
end program test
