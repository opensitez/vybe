! vybe-test: fortran/intent_optional_extended/default_power_exponent_two
! origin: languages/fortran/tests/fortran/test_intent_optional_extended.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(2) = [ 25, 125 ]
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 2) then
    print *, "FAIL: more than 2 line(s)"
    stop 1
end if
if ((opt_pow(5)) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", opt_pow(5), "]"
    stop 1
end if
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 2) then
    print *, "FAIL: more than 2 line(s)"
    stop 1
end if
if ((opt_pow(5, 3)) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", opt_pow(5, 3), "]"
    stop 1
end if
if (vybe_check_i /= 2) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 2"
    stop 1
end if
contains
integer function opt_pow(base, exp)
integer, intent(in) :: base
integer, intent(in), optional :: exp
integer :: use_e, i
if (present(exp)) then
use_e = exp
else
use_e = 2
end if
opt_pow = 1
do i = 1, use_e
opt_pow = opt_pow * base
end do
end function opt_pow
end program t
