! vybe-test: fortran/character_scan_verify_extended/repeat_single_char_five
! origin: languages/fortran/tests/fortran/test_character_scan_verify_extended.rs
program t
if (trim(repeat('x', 5)) /= "xxxxx") then
    print *, "FAIL: want [xxxxx] got [", repeat('x', 5), "]"
    stop 1
end if
end program t
