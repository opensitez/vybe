! vybe-test: fortran/character_scan_verify_extended/scan_forward_first_hex_in_uuid
! origin: languages/fortran/tests/fortran/test_character_scan_verify_extended.rs
program t
character(len=8) :: s = '019af2b0'
if ((scan(s, 'abcdef')) /= 3) then
    print *, "FAIL: want [3] got [", scan(s, 'abcdef'), "]"
    stop 1
end if
end program t
