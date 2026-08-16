! vybe-test: fortran/intent_optional_extended/default_repeat_count_one
! origin: languages/fortran/tests/fortran/test_intent_optional_extended.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(2) = [ 9, 27 ]
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 2) then
    print *, "FAIL: more than 2 line(s)"
    stop 1
end if
if ((repeat_val(9)) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", repeat_val(9), "]"
    stop 1
end if
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 2) then
    print *, "FAIL: more than 2 line(s)"
    stop 1
end if
if ((repeat_val(9, 3)) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", repeat_val(9, 3), "]"
    stop 1
end if
if (vybe_check_i /= 2) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 2"
    stop 1
end if
contains
integer function repeat_val(x, n)
integer, intent(in) :: x
integer, intent(in), optional :: n
integer :: use_n, i
if (present(n)) then
use_n = n
else
use_n = 1
end if
repeat_val = 0
do i = 1, use_n
repeat_val = repeat_val + x
end do
end function repeat_val
end program t
