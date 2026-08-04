! vybe-test: fortran/character_scan_verify_extended/ichar_digit_1
! origin: languages/fortran/tests/fortran/test_character_scan_verify_extended.rs
program t
if ((ichar('1')) /= 49) then
    print *, "FAIL: want [49] got [", ichar('1'), "]"
    stop 1
end if
end program t
