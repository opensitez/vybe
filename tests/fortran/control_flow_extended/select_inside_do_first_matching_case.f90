! vybe-test: fortran/control_flow_extended/select_inside_do_first_matching_case
! origin: languages/fortran/tests/fortran/test_control_flow_extended.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 15 ]
integer :: i, hits
hits = 0
do i = 1, 6
select case (mod(i, 3))
case (0)
hits = hits + 10
case (1)
hits = hits + 1
case (2)
hits = hits + 2
end select
end do
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if ((hits) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", hits, "]"
    stop 1
end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program t
