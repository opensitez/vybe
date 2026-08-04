! vybe-test: fortran/character_intrinsics_extended/if_verify_branch_detects_symbol
! origin: languages/fortran/tests/fortran/test_character_intrinsics_extended.rs
program t
integer :: vybe_check_i = 0
character(len=5) :: vybe_check_w(1) = [ "other" ]
character(len=5) :: s = 'safe1'
if (verify(s, 'abcdefghijklmnopqrstuvwxyz') == 0) then
    vybe_check_i = vybe_check_i + 1
  if (vybe_check_i > 1) then
      print *, "FAIL: more than 1 line(s)"
      stop 1
  end if
  if (trim('letters') /= trim(vybe_check_w(vybe_check_i))) then
      print *, "FAIL at ", vybe_check_i, " got [", 'letters', "]"
      stop 1
  end if
else
    vybe_check_i = vybe_check_i + 1
  if (vybe_check_i > 1) then
      print *, "FAIL: more than 1 line(s)"
      stop 1
  end if
  if (trim('other') /= trim(vybe_check_w(vybe_check_i))) then
      print *, "FAIL at ", vybe_check_i, " got [", 'other', "]"
      stop 1
  end if
end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program t
