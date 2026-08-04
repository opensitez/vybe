! vybe-test: fortran/control_flow_extended/nested_labeled_do_counts_nine
! origin: languages/fortran/tests/fortran/test_control_flow_extended.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 9 ]
integer :: i, j, c
c = 0
do 200 i = 1, 3
do 100 j = 1, 3
c = c + 1
100 continue
200 continue
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
