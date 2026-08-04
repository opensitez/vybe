! vybe-test: fortran/character_scan_verify_extended/adjustl_moves_leading_blanks
! origin: languages/fortran/tests/fortran/test_character_scan_verify_extended.rs
program t
character(len=10) :: s = '   data   '
if (trim(trim(adjustl(s))) /= "data") then
    print *, "FAIL: want [data] got [", trim(adjustl(s)), "]"
    stop 1
end if
end program t
