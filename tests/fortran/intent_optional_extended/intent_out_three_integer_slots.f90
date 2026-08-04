! vybe-test: fortran/intent_optional_extended/intent_out_three_integer_slots
! origin: languages/fortran/tests/fortran/test_intent_optional_extended.rs
program t
integer :: a, b, c
call triple_out(a, b, c)
if ((a + b + c) /= 9) then
    print *, "FAIL: want [9] got [", a + b + c, "]"
    stop 1
end if
contains
subroutine triple_out(x, y, z)
integer, intent(out) :: x, y, z
x = 2
y = 3
z = 4
end subroutine triple_out
end program t
