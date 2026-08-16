! vybe-test: fortran/intent_attributes/intent_attributes_11
! origin: languages/fortran/tests/fortran/test_intent_attributes.rs
program t
integer :: total
call s(4)
if (total /= 4) then
    print *, "FAIL: want [4] got [", total, "]"
    stop 1
end if
call s(4, 6)
if (total /= 10) then
    print *, "FAIL: want [10] got [", total, "]"
    stop 1
end if
contains
subroutine s(x, y)
integer, intent(in) :: x
integer, optional, intent(in) :: y
total = x
if (present(y)) total = total + y
end subroutine s
end program t
