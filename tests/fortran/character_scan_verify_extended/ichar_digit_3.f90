! vybe-test: fortran/character_scan_verify_extended/ichar_digit_3
! origin: languages/fortran/tests/fortran/test_character_scan_verify_extended.rs
program t
if ((ichar('3')) /= 51) then
    print *, "FAIL: want [51] got [", ichar('3'), "]"
    stop 1
end if
end program t
