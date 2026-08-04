! vybe-test: fortran/intent_optional_extended/present_flag_reports_zero_when_absent
! origin: languages/fortran/tests/fortran/test_intent_optional_extended.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 0 ]
call show_present(4)
contains
subroutine show_present(x, y)
integer, intent(in) :: x
integer, intent(in), optional :: y
if (present(y)) then
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if ((1) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", 1, "]"
    stop 1
end if
else
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if ((0) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", 0, "]"
    stop 1
end if
end if
end subroutine show_present
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program t
