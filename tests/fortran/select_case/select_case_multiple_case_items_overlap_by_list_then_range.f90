! vybe-test: fortran/select_case/select_case_multiple_case_items_overlap_by_list_then_range
! origin: languages/fortran/tests/fortran/test_select_case.rs

program t
integer :: vybe_check_i = 0
character(len=4) :: vybe_check_w(1) = [ "list" ]
integer :: n
n = 4
select case (n)
case (3, 4, 5)
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if (trim('list') /= trim(vybe_check_w(vybe_check_i))) then
    print *, "FAIL at ", vybe_check_i, " got [", 'list', "]"
    stop 1
end if
case (1:10)
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if (trim('range') /= trim(vybe_check_w(vybe_check_i))) then
    print *, "FAIL at ", vybe_check_i, " got [", 'range', "]"
    stop 1
end if
end select
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program t
