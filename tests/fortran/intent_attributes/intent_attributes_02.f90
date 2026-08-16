! vybe-test: fortran/intent_attributes/intent_attributes_02
! origin: languages/fortran/tests/fortran/test_intent_attributes.rs
program t
integer :: v
v = 99
call s(v)
if (v /= 5) then
    print *, "FAIL: want [5] got [", v, "]"
    stop 1
end if
contains
subroutine s(x)
integer, intent(out) :: x
x = 5
end subroutine s
end program t
