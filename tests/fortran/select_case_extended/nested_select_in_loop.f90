! vybe-test: fortran/select_case_extended/nested_select_in_loop
! origin: languages/fortran/tests/fortran/test_select_case_extended.rs
program t
integer :: vybe_check_i = 0
character(len=11) :: vybe_check_w(4) = [ "one-one", "one-two", "outer range", "outer range" ]
integer :: i, j
do i = 1, 3
select case (i)
case (1)
do j = 1, 2
select case (j)
case (1)
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 4) then
    print *, "FAIL: more than 4 line(s)"
    stop 1
end if
if (trim("one-one") /= trim(vybe_check_w(vybe_check_i))) then
    print *, "FAIL at ", vybe_check_i, " got [", "one-one", "]"
    stop 1
end if
case (2)
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 4) then
    print *, "FAIL: more than 4 line(s)"
    stop 1
end if
if (trim("one-two") /= trim(vybe_check_w(vybe_check_i))) then
    print *, "FAIL at ", vybe_check_i, " got [", "one-two", "]"
    stop 1
end if
end select
end do
case (2:3)
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 4) then
    print *, "FAIL: more than 4 line(s)"
    stop 1
end if
if (trim("outer range") /= trim(vybe_check_w(vybe_check_i))) then
    print *, "FAIL at ", vybe_check_i, " got [", "outer range", "]"
    stop 1
end if
end select
end do
if (vybe_check_i /= 4) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 4"
    stop 1
end if
end program t
