! vybe-test: fortran/alternate_returns/alternate_returns_runtime_second_alternate_is_taken_on_return_2
! origin: languages/fortran/tests/fortran/test_alternate_returns.rs
program p
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 3 ]
integer :: x
x = 0
call s(*10,*20)
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if ((x) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", x, "]"
    stop 1
end if
10 x = 2
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if ((x) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", x, "]"
    stop 1
end if
20 x = 3
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if ((x) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", x, "]"
    stop 1
end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program p
subroutine s(*,*)
return 2
end
