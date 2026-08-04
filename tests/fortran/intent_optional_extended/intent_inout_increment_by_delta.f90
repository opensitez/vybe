! vybe-test: fortran/intent_optional_extended/intent_inout_increment_by_delta
! origin: languages/fortran/tests/fortran/test_intent_optional_extended.rs
program t
integer :: n
n = 10
call bump_by(n, 7)
if ((n) /= 17) then
    print *, "FAIL: want [17] got [", n, "]"
    stop 1
end if
contains
subroutine bump_by(x, delta)
integer, intent(inout) :: x
integer, intent(in) :: delta
x = x + delta
end subroutine bump_by
end program t
