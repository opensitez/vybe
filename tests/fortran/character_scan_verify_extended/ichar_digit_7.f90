! vybe-test: fortran/character_scan_verify_extended/ichar_digit_7
! origin: languages/fortran/tests/fortran/test_character_scan_verify_extended.rs
program t
if ((ichar('7')) /= 55) then
    print *, "FAIL: want [55] got [", ichar('7'), "]"
    stop 1
end if
end program t
