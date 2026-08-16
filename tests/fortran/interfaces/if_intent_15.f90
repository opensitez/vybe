! vybe-test: fortran/interfaces/if_intent_15
! origin: languages/fortran/tests/fortran/test_interfaces.rs
program t
integer :: seen
seen = 0
call s(7)
if (seen /= 7) then
    print *, "FAIL: want [7] got [", seen, "]"
    stop 1
end if
contains
subroutine s(x)
integer,intent(in)::x
seen = x
end subroutine s
end program t
