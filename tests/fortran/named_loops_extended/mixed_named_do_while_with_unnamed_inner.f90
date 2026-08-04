! vybe-test: fortran/named_loops_extended/mixed_named_do_while_with_unnamed_inner
! origin: languages/fortran/tests/fortran/test_named_loops_extended.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 5 ]
integer :: n, j, c
n = 0
c = 0
spin: do while (n < 5)
n = n + 1
do j = 1, 3
if (j == 2) cycle spin
c = c + 1
end do
end do spin
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if ((c) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", c, "]"
    stop 1
end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program t
