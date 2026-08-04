! vybe-test: fortran/character_scan_verify_extended/adjustr_right_aligns_in_field
! origin: languages/fortran/tests/fortran/test_character_scan_verify_extended.rs
program t
character(len=6) :: s = 'xy    '
if (trim(trim(adjustr(s))) /= "xy") then
    print *, "FAIL: want [xy] got [", trim(adjustr(s)), "]"
    stop 1
end if
end program t
