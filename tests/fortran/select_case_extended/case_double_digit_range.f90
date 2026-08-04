! vybe-test: fortran/select_case_extended/case_double_digit_range
! origin: languages/fortran/tests/fortran/test_select_case_extended.rs
program t
integer :: vybe_check_i = 0
character(len=12) :: vybe_check_w(1) = [ "double digit" ]
integer :: n = 50
select case (n)
case (0:9)
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if (trim("single digit") /= trim(vybe_check_w(vybe_check_i))) then
    print *, "FAIL at ", vybe_check_i, " got [", "single digit", "]"
    stop 1
end if
case (10:99)
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if (trim("double digit") /= trim(vybe_check_w(vybe_check_i))) then
    print *, "FAIL at ", vybe_check_i, " got [", "double digit", "]"
    stop 1
end if
case (100:999)
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if (trim("triple digit or more") /= trim(vybe_check_w(vybe_check_i))) then
    print *, "FAIL at ", vybe_check_i, " got [", "triple digit or more", "]"
    stop 1
end if
end select
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program t
