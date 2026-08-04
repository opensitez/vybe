! vybe-test: fortran/intent_optional_extended/intent_inout_reverse_two_element
! origin: languages/fortran/tests/fortran/test_intent_optional_extended.rs
program t
integer :: pair(2)
pair = [3, 9]
call reverse2(pair)
if ((pair(1)) /= 9) then
    print *, "FAIL: want [9] got [", pair(1), "]"
    stop 1
end if
if ((pair(2)) /= 3) then
    print *, "FAIL: want [3] got [", pair(2), "]"
    stop 1
end if
contains
subroutine reverse2(v)
integer, intent(inout) :: v(2)
integer :: t
t = v(1)
v(1) = v(2)
v(2) = t
end subroutine reverse2
end program t
