! vybe-test: fortran/intent_attributes/intent_attributes_12
! origin: languages/fortran/tests/fortran/test_intent_attributes.rs
program t
integer :: seen
call s(7)
if (seen /= 8) then
    print *, "FAIL: want [8] got [", seen, "]"
    stop 1
end if
contains
subroutine s(x)
integer, intent(in) :: x
integer :: y
y = x + 1
seen = y
end subroutine s
end program t
