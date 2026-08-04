! vybe-test: fortran/character_scan_verify_extended/adjustr_moves_trailing_content
! origin: languages/fortran/tests/fortran/test_character_scan_verify_extended.rs
program t
character(len=8) :: s = 'code    '
if (trim(trim(adjustr(s))) /= "code") then
    print *, "FAIL: want [code] got [", trim(adjustr(s)), "]"
    stop 1
end if
end program t
