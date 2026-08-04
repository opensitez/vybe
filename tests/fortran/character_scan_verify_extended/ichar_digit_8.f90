! vybe-test: fortran/character_scan_verify_extended/ichar_digit_8
! origin: languages/fortran/tests/fortran/test_character_scan_verify_extended.rs
program t
if ((ichar('8')) /= 56) then
    print *, "FAIL: want [56] got [", ichar('8'), "]"
    stop 1
end if
end program t
