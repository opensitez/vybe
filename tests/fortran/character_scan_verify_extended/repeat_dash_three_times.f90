! vybe-test: fortran/character_scan_verify_extended/repeat_dash_three_times
! origin: languages/fortran/tests/fortran/test_character_scan_verify_extended.rs
program t
if (trim(repeat('ab', 3)) /= "ababab") then
    print *, "FAIL: want [ababab] got [", repeat('ab', 3), "]"
    stop 1
end if
end program t
