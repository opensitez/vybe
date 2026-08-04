! vybe-test: fortran/control/ctrl_do_concurrent_17
! origin: languages/fortran/tests/fortran/test_control.rs
program p
integer :: vybe_check_i = 0
integer :: vybe_check_w(3) = [ 1, 2, 3 ]
integer::i
do concurrent (i=1:3)
  vybe_check_i = vybe_check_i + 1
 if (vybe_check_i > 3) then
     print *, "FAIL: more than 3 line(s)"
     stop 1
 end if
 if ((i) /= vybe_check_w(vybe_check_i)) then
     print *, "FAIL at ", vybe_check_i, " got [", i, "]"
     stop 1
 end if
end do
if (vybe_check_i /= 3) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 3"
    stop 1
end if
end program p
