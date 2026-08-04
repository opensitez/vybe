! vybe-test: fortran/character_scan_verify_extended/ichar_digit_4
! origin: languages/fortran/tests/fortran/test_character_scan_verify_extended.rs
program t
if ((ichar('4')) /= 52) then
    print *, "FAIL: want [52] got [", ichar('4'), "]"
    stop 1
end if
end program t
