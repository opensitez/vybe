! vybe-test: fortran/intent_attributes/intent_attributes_07
! origin: languages/fortran/tests/fortran/test_intent_attributes.rs
program t
integer :: seen
seen = 0
call s("hello")
if (seen /= 5) then
    print *, "FAIL: want [5] got [", seen, "]"
    stop 1
end if
contains
subroutine s(x)
character(len=*), intent(in) :: x
seen = len(x)
end subroutine s
end program t
