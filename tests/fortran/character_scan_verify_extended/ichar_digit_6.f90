! vybe-test: fortran/character_scan_verify_extended/ichar_digit_6
! origin: languages/fortran/tests/fortran/test_character_scan_verify_extended.rs
program t
if ((ichar('6')) /= 54) then
    print *, "FAIL: want [54] got [", ichar('6'), "]"
    stop 1
end if
end program t
