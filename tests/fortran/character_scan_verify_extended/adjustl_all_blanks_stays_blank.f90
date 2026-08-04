! vybe-test: fortran/character_scan_verify_extended/adjustl_all_blanks_stays_blank
! origin: languages/fortran/tests/fortran/test_character_scan_verify_extended.rs
program t
character(len=5) :: s = '     '
if (trim(trim(adjustl(s))) /= "") then
    print *, "FAIL: want [] got [", trim(adjustl(s)), "]"
    stop 1
end if
end program t
