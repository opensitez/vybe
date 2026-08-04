! vybe-test: fortran/intent_optional_extended/optional_logical_invert_when_present
! origin: languages/fortran/tests/fortran/test_intent_optional_extended.rs
program t
if ((maybe_not(.true.)) .neqv. .true.) then
    print *, "FAIL: want [true] got [", maybe_not(.true.), "]"
    stop 1
end if
if ((maybe_not(.true., .true.)) .neqv. .false.) then
    print *, "FAIL: want [false] got [", maybe_not(.true., .true.), "]"
    stop 1
end if
contains
logical function maybe_not(v, flip)
logical, intent(in) :: v
logical, intent(in), optional :: flip
if (present(flip)) then
maybe_not = .not. v
else
maybe_not = v
end if
end function maybe_not
end program t
