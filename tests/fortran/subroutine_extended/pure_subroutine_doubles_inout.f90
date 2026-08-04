! vybe-test: fortran/subroutine_extended/pure_subroutine_doubles_inout
! origin: languages/fortran/tests/fortran/test_subroutine_extended.rs
program t
integer :: n
n = 11
call pdbl(n)
if ((n) /= 22) then
    print *, "FAIL: want [22] got [", n, "]"
    stop 1
end if
contains
pure subroutine pdbl(x)
integer, intent(inout) :: x
x = x * 2
end subroutine pdbl
end program t
