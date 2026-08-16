! vybe-test: fortran/intent_attributes/intent_attributes_10
! origin: languages/fortran/tests/fortran/test_intent_attributes.rs
program t
integer :: seen
call s()
if (seen /= -1) then
    print *, "FAIL: want [-1] got [", seen, "]"
    stop 1
end if
call s(3)
if (seen /= 3) then
    print *, "FAIL: want [3] got [", seen, "]"
    stop 1
end if
contains
subroutine s(x)
integer, optional, intent(in) :: x
if (present(x)) then
    seen = x
else
    seen = -1
end if
end subroutine s
end program t
