! vybe-test: fortran/intent_optional_extended/present_drives_merge_style_default
! origin: languages/fortran/tests/fortran/test_intent_optional_extended.rs
program t
if ((pick_or(7, 3)) /= 3) then
    print *, "FAIL: want [3] got [", pick_or(7, 3), "]"
    stop 1
end if
if ((pick_or(7)) /= 7) then
    print *, "FAIL: want [7] got [", pick_or(7), "]"
    stop 1
end if
contains
integer function pick_or(base, alt)
integer, intent(in) :: base
integer, intent(in), optional :: alt
if (present(alt)) then
pick_or = alt
else
pick_or = base
end if
end function pick_or
end program t
