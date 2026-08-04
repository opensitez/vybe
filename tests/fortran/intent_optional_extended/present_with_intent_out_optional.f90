! vybe-test: fortran/intent_optional_extended/present_with_intent_out_optional
! origin: languages/fortran/tests/fortran/test_intent_optional_extended.rs
program t
integer :: r
call maybe_fill(5, r)
if ((r) /= 5) then
    print *, "FAIL: want [5] got [", r, "]"
    stop 1
end if
contains
subroutine maybe_fill(seed, outv, use_out)
integer, intent(in) :: seed
integer, intent(out) :: outv
logical, intent(in), optional :: use_out
if (present(use_out) .and. use_out) then
outv = seed * 2
else
outv = seed
end if
end subroutine maybe_fill
end program t
