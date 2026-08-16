! vybe-test: fortran/interfaces/if_optional_two_37
! origin: languages/fortran/tests/fortran/test_interfaces.rs
program t
integer :: n
call s()
if (n /= 0) then
    print *, "FAIL: want [0] got [", n, "]"
    stop 1
end if
call s(1)
if (n /= 1) then
    print *, "FAIL: want [1] got [", n, "]"
    stop 1
end if
call s(y=2)
if (n /= 1) then
    print *, "FAIL: want [1] got [", n, "]"
    stop 1
end if
call s(1, 2)
if (n /= 2) then
    print *, "FAIL: want [2] got [", n, "]"
    stop 1
end if
contains
subroutine s(x,y)
integer, optional :: x,y
n = 0
if (present(x)) n = n + 1
if (present(y)) n = n + 1
end subroutine s
end program t
