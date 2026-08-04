! vybe-test: fortran/intent_optional_extended/optional_pair_subroutine_flags
! origin: languages/fortran/tests/fortran/test_intent_optional_extended.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(3) = [ 3, 5, 6 ]
call report_flags(3)
call report_flags(3, 2)
call report_flags(3, 2, 1)
contains
subroutine report_flags(a, b, c)
integer, intent(in) :: a
integer, intent(in), optional :: b, c
integer :: s
s = a
if (present(b)) s = s + b
if (present(c)) s = s + c
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 3) then
    print *, "FAIL: more than 3 line(s)"
    stop 1
end if
if ((s) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", s, "]"
    stop 1
end if
end subroutine report_flags
if (vybe_check_i /= 3) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 3"
    stop 1
end if
end program t
