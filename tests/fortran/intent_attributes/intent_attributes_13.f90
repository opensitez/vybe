! vybe-test: fortran/intent_attributes/intent_attributes_13
! origin: languages/fortran/tests/fortran/test_intent_attributes.rs
program t
integer :: a, b, c
a = 3
b = 4
c = 0
call s(a, b, c)
if (c /= 7) then
    print *, "FAIL: want [7] got [", c, "]"
    stop 1
end if
if (a /= 7) then
    print *, "FAIL: want [7] got [", a, "]"
    stop 1
end if
contains
subroutine s(a, b, c)
integer, intent(inout) :: a
integer, intent(in) :: b
integer, intent(out) :: c
c = a + b
a = c
end subroutine s
end program t
