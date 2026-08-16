! vybe-test: fortran/intent_attributes/intent_attributes_01
! origin: languages/fortran/tests/fortran/test_intent_attributes.rs
program t
integer :: seen
seen = 0
call s(4)
if (seen /= 4) then
    print *, "FAIL: want [4] got [", seen, "]"
    stop 1
end if
contains
subroutine s(x)
integer, intent(in) :: x
seen = x
end subroutine s
end program t
