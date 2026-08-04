! vybe-test: fortran/select_case_extended/case_multi_values_loop_match
! origin: languages/fortran/tests/fortran/test_select_case_extended.rs
program t
integer :: vybe_check_i = 0
character(len=5) :: vybe_check_w(6) = [ "match", "match", "no", "no", "no", "match" ]
integer :: i
do i = 1, 6
select case (i)
case (1, 2, 6)
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 6) then
    print *, "FAIL: more than 6 line(s)"
    stop 1
end if
if (trim("match") /= trim(vybe_check_w(vybe_check_i))) then
    print *, "FAIL at ", vybe_check_i, " got [", "match", "]"
    stop 1
end if
case default
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 6) then
    print *, "FAIL: more than 6 line(s)"
    stop 1
end if
if (trim("no") /= trim(vybe_check_w(vybe_check_i))) then
    print *, "FAIL at ", vybe_check_i, " got [", "no", "]"
    stop 1
end if
end select
end do
if (vybe_check_i /= 6) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 6"
    stop 1
end if
end program t
