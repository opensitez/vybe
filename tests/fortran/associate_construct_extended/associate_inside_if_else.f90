! vybe-test: fortran/associate_construct_extended/associate_inside_if_else
! origin: languages/fortran/tests/fortran/test_associate_construct_extended.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 4 ]
integer :: n = -4
if (n >= 0) then
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if ((0) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", 0, "]"
    stop 1
end if
else
associate (mag => abs(n))
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if ((mag) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", mag, "]"
    stop 1
end if
end associate
end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program t
