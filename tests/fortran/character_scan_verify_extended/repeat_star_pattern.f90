! vybe-test: fortran/character_scan_verify_extended/repeat_star_pattern
! origin: languages/fortran/tests/fortran/test_character_scan_verify_extended.rs
program t
if (trim(repeat('*', 4)) /= "****") then
    print *, "FAIL: want [****] got [", repeat('*', 4), "]"
    stop 1
end if
end program t
