! vybe-test: fortran/character_scan_verify_extended/ichar_digit_0
! origin: languages/fortran/tests/fortran/test_character_scan_verify_extended.rs
program t
if ((ichar('0')) /= 48) then
    print *, "FAIL: want [48] got [", ichar('0'), "]"
    stop 1
end if
end program t
