! vybe-test: fortran/select_case_extended/case_lower_open_range_match
! origin: languages/fortran/tests/fortran/test_select_case_extended.rs
program t
integer :: vybe_check_i = 0
! n = -99 falls in `case (-99:0)`, whose branch asserts "small". The expected
! word was "too small" — the label on the branch ABOVE, which never runs — so
! gfortran failed this test too.
character(len=9) :: vybe_check_w(1) = [ "small" ]
integer :: n
n = -99
select case (n)
case (-200:-100)
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if (trim("too small") /= trim(vybe_check_w(vybe_check_i))) then
    print *, "FAIL at ", vybe_check_i, " got [", "too small", "]"
    stop 1
end if
case (-99:0)
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if (trim("small") /= trim(vybe_check_w(vybe_check_i))) then
    print *, "FAIL at ", vybe_check_i, " got [", "small", "]"
    stop 1
end if
case (1:200)
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if (trim("large") /= trim(vybe_check_w(vybe_check_i))) then
    print *, "FAIL at ", vybe_check_i, " got [", "large", "]"
    stop 1
end if
end select
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program t
