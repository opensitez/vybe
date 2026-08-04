! vybe-test: fortran/character_scan_verify_extended/ichar_digit_9
! origin: languages/fortran/tests/fortran/test_character_scan_verify_extended.rs
program t
if ((ichar('9')) /= 57) then
    print *, "FAIL: want [57] got [", ichar('9'), "]"
    stop 1
end if
end program t
