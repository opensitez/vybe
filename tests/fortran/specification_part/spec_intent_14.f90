! vybe-test: fortran/specification_part/spec_intent_14
! origin: languages/fortran/tests/fortran/test_specification_part.rs
program t
implicit none
integer :: seen
seen = 0
call s(3)
if (seen /= 3) then
    print *, "FAIL: want [3] got [", seen, "]"
    stop 1
end if
contains
subroutine s(x)
implicit none
integer, intent(in) :: x
seen = x
end subroutine s
end program t
