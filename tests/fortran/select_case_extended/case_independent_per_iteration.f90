! vybe-test: fortran/select_case_extended/case_independent_per_iteration
! origin: languages/fortran/tests/fortran/test_select_case_extended.rs
program t
integer :: vybe_check_i = 0
character(len=5) :: vybe_check_w(4) = [ "alpha", "beta", "gamma", "omega" ]
integer :: i
do i = 1, 4
select case (i)
case (1)
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 4) then
    print *, "FAIL: more than 4 line(s)"
    stop 1
end if
if (trim("alpha") /= trim(vybe_check_w(vybe_check_i))) then
    print *, "FAIL at ", vybe_check_i, " got [", "alpha", "]"
    stop 1
end if
case (2)
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 4) then
    print *, "FAIL: more than 4 line(s)"
    stop 1
end if
if (trim("beta") /= trim(vybe_check_w(vybe_check_i))) then
    print *, "FAIL at ", vybe_check_i, " got [", "beta", "]"
    stop 1
end if
case (3)
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 4) then
    print *, "FAIL: more than 4 line(s)"
    stop 1
end if
if (trim("gamma") /= trim(vybe_check_w(vybe_check_i))) then
    print *, "FAIL at ", vybe_check_i, " got [", "gamma", "]"
    stop 1
end if
case default
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 4) then
    print *, "FAIL: more than 4 line(s)"
    stop 1
end if
if (trim("omega") /= trim(vybe_check_w(vybe_check_i))) then
    print *, "FAIL at ", vybe_check_i, " got [", "omega", "]"
    stop 1
end if
end select
end do
if (vybe_check_i /= 4) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 4"
    stop 1
end if
end program t
