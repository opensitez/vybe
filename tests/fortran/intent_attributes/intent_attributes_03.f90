! vybe-test: fortran/intent_attributes/intent_attributes_03
! origin: languages/fortran/tests/fortran/test_intent_attributes.rs
program t
integer :: v
v = 4
call s(v)
if (v /= 9) then
    print *, "FAIL: want [9] got [", v, "]"
    stop 1
end if
contains
subroutine s(x)
integer, intent(inout) :: x
x = x + 5
end subroutine s
end program t
