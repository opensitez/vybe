! vybe-test: fortran/integer_mod_division/gcd_loop_mod_99_and_78
! origin: languages/fortran/tests/fortran/test_integer_mod_division.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 3 ]
integer :: a, b, tmp
a = 99
b = 78
do while (b /= 0)
tmp = b
b = mod(a, b)
a = tmp
end do
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if ((a) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", a, "]"
    stop 1
end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program t
