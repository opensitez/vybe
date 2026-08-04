! vybe-test: fortran/if_blocks/if_logical_chain_true_mid_branch
! origin: languages/fortran/tests/fortran/test_if_blocks.rs
program t
integer :: vybe_check_i = 0
character(len=3) :: vybe_check_w(1) = [ "pos" ]
integer :: x
logical :: ok
x = 5
ok = .false.
if (x < 0) then
    vybe_check_i = vybe_check_i + 1
  if (vybe_check_i > 1) then
      print *, "FAIL: more than 1 line(s)"
      stop 1
  end if
  if (trim("neg") /= trim(vybe_check_w(vybe_check_i))) then
      print *, "FAIL at ", vybe_check_i, " got [", "neg", "]"
      stop 1
  end if
else if (x > 0 .and. .not. ok) then
    vybe_check_i = vybe_check_i + 1
  if (vybe_check_i > 1) then
      print *, "FAIL: more than 1 line(s)"
      stop 1
  end if
  if (trim("pos") /= trim(vybe_check_w(vybe_check_i))) then
      print *, "FAIL at ", vybe_check_i, " got [", "pos", "]"
      stop 1
  end if
else if (x == 5) then
    vybe_check_i = vybe_check_i + 1
  if (vybe_check_i > 1) then
      print *, "FAIL: more than 1 line(s)"
      stop 1
  end if
  if (trim("five") /= trim(vybe_check_w(vybe_check_i))) then
      print *, "FAIL at ", vybe_check_i, " got [", "five", "]"
      stop 1
  end if
else
    vybe_check_i = vybe_check_i + 1
  if (vybe_check_i > 1) then
      print *, "FAIL: more than 1 line(s)"
      stop 1
  end if
  if (trim("other") /= trim(vybe_check_w(vybe_check_i))) then
      print *, "FAIL at ", vybe_check_i, " got [", "other", "]"
      stop 1
  end if
end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if

