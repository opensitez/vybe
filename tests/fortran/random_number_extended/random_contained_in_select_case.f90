! vybe-test: fortran/random_number_extended/random_contained_in_select_case
! origin: languages/fortran/tests/fortran/test_random_number_extended.rs
program t
integer :: vybe_check_i = 0
character(len=2) :: vybe_check_w(1) = [ "in" ]
real :: r
call random_number(r)
select case (merge(1, 0, r < 1.0))
case (1)
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if (trim('in') /= trim(vybe_check_w(vybe_check_i))) then
    print *, "FAIL at ", vybe_check_i, " got [", 'in', "]"
    stop 1
end if
case default
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if (trim('out') /= trim(vybe_check_w(vybe_check_i))) then
    print *, "FAIL at ", vybe_check_i, " got [", 'out', "]"
    stop 1
end if
end select
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program t
