! vybe-test: fortran/intent_optional_extended/optional_offset_defaults_to_zero
! origin: languages/fortran/tests/fortran/test_intent_optional_extended.rs
program t
if ((shift_val(8)) /= 8) then
    print *, "FAIL: want [8] got [", shift_val(8), "]"
    stop 1
end if
if ((shift_val(8, 3)) /= 11) then
    print *, "FAIL: want [11] got [", shift_val(8, 3), "]"
    stop 1
end if
contains
integer function shift_val(x, off)
integer, intent(in) :: x
integer, intent(in), optional :: off
if (present(off)) then
shift_val = x + off
else
shift_val = x
end if
end function shift_val
end program t
