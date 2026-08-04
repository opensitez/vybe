! vybe-test: fortran/character_scan_verify_extended/scan_backward_last_lower_in_caps
! origin: languages/fortran/tests/fortran/test_character_scan_verify_extended.rs
program t
character(len=6) :: s = 'ABCdEF'
if ((scan(s, 'abcdefghijklmnopqrstuvwxyz', .true.)) /= 4) then
    print *, "FAIL: want [4] got [", scan(s, 'abcdefghijklmnopqrstuvwxyz', .true.), "]"
    stop 1
end if
end program t
