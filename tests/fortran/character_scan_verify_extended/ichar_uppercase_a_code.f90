! vybe-test: fortran/character_scan_verify_extended/ichar_uppercase_a_code
! origin: languages/fortran/tests/fortran/test_character_scan_verify_extended.rs
program t
if ((ichar('A')) /= 65) then
    print *, "FAIL: want [65] got [", ichar('A'), "]"
    stop 1
end if
end program t
