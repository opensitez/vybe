! vybe-test: fortran/subroutine_extended/optional_subroutine_extra_term
! origin: languages/fortran/tests/fortran/test_subroutine_extended.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(2) = [ 4, 10 ]
call accumulate(4)
call accumulate(4, 6)
contains
subroutine accumulate(base, extra)
integer, intent(in) :: base
integer, intent(in), optional :: extra
integer :: total
total = base
if (present(extra)) total = total + extra
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 2) then
    print *, "FAIL: more than 2 line(s)"
    stop 1
end if
if ((total) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", total, "]"
    stop 1
end if
end subroutine accumulate
if (vybe_check_i /= 2) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 2"
    stop 1
end if
end program t
