! vybe-test: fortran/character_scan_verify_extended/repeat_zero_is_empty
! origin: languages/fortran/tests/fortran/test_character_scan_verify_extended.rs
program t
if (trim(repeat('hi', 0)) /= "") then
    print *, "FAIL: want [] got [", repeat('hi', 0), "]"
    stop 1
end if
end program t
