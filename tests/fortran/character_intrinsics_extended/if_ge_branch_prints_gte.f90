! vybe-test: fortran/character_intrinsics_extended/if_ge_branch_prints_gte
! origin: languages/fortran/tests/fortran/test_character_intrinsics_extended.rs
program t
integer :: vybe_check_i = 0
character(len=3) :: vybe_check_w(1) = [ "gte" ]
if ('zebra' >= 'apple') then
    vybe_check_i = vybe_check_i + 1
  if (vybe_check_i > 1) then
      print *, "FAIL: more than 1 line(s)"
      stop 1
  end if
  if (trim('gte') /= trim(vybe_check_w(vybe_check_i))) then
      print *, "FAIL at ", vybe_check_i, " got [", 'gte', "]"
      stop 1
  end if
else
    vybe_check_i = vybe_check_i + 1
  if (vybe_check_i > 1) then
      print *, "FAIL: more than 1 line(s)"
      stop 1
  end if
  if (trim('lt') /= trim(vybe_check_w(vybe_check_i))) then
      print *, "FAIL at ", vybe_check_i, " got [", 'lt', "]"
      stop 1
  end if
end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program t
