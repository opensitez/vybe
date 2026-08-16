! vybe-test: fortran/interfaces/if_optional_06
! origin: languages/fortran/tests/fortran/test_interfaces.rs
program t
integer :: seen
seen = -1
call s()
if (seen /= 0) then
    print *, "FAIL: want [0] got [", seen, "]"
    stop 1
end if
call s(5)
if (seen /= 5) then
    print *, "FAIL: want [5] got [", seen, "]"
    stop 1
end if
contains
subroutine s(x)
integer, optional :: x
if (present(x)) then
    seen = x
else
    seen = 0
end if
end subroutine s
end program t
