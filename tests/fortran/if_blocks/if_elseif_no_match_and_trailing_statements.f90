! vybe-test: fortran/if_blocks/if_elseif_no_match_and_trailing_statements
! origin: languages/fortran/tests/fortran/test_if_blocks.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 9 ]
integer :: x
x = 99
if (x == 1) then
    vybe_check_i = vybe_check_i + 1
  if (vybe_check_i > 1) then
      print *, "FAIL: more than 1 line(s)"
      stop 1
  end if
  if ((1) /= vybe_check_w(vybe_check_i)) then
      print *, "FAIL at ", vybe_check_i, " got [", 1, "]"
      stop 1
  end if
else if (x == 2) then
    vybe_check_i = vybe_check_i + 1
  if (vybe_check_i > 1) then
      print *, "FAIL: more than 1 line(s)"
      stop 1
  end if
  if ((2) /= vybe_check_w(vybe_check_i)) then
      print *, "FAIL at ", vybe_check_i, " got [", 2, "]"
      stop 1
  end if
else if (x == 3) then
    vybe_check_i = vybe_check_i + 1
  if (vybe_check_i > 1) then
      print *, "FAIL: more than 1 line(s)"
      stop 1
  end if
  if ((3) /= vybe_check_w(vybe_check_i)) then
      print *, "FAIL at ", vybe_check_i, " got [", 3, "]"
      stop 1
  end if
end if
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if ((9) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", 9, "]"
    stop 1
end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if

end program t
