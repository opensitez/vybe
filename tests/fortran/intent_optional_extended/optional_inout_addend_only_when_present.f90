! vybe-test: fortran/intent_optional_extended/optional_inout_addend_only_when_present
! origin: languages/fortran/tests/fortran/test_intent_optional_extended.rs
program t
integer :: n
n = 4
call maybe_add(n)
if ((n) /= 4) then
    print *, "FAIL: want [4] got [", n, "]"
    stop 1
end if
contains
subroutine maybe_add(x, extra)
integer, intent(inout) :: x
integer, intent(in), optional :: extra
if (present(extra)) x = x + extra
end subroutine maybe_add
end program t
