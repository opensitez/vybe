! vybe-test: fortran/integer_mod_division/do_while_mod_condition_counts_steps
! origin: languages/fortran/tests/fortran/test_integer_mod_division.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 6 ]
integer :: n, steps
n = 100
steps = 0
do while (n > 1)
n = n / 2
steps = steps + 1
end do
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if ((steps) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", steps, "]"
    stop 1
end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program t
