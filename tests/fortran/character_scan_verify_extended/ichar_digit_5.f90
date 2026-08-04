! vybe-test: fortran/character_scan_verify_extended/ichar_digit_5
! origin: languages/fortran/tests/fortran/test_character_scan_verify_extended.rs
program t
if ((ichar('5')) /= 53) then
    print *, "FAIL: want [53] got [", ichar('5'), "]"
    stop 1
end if
end program t
