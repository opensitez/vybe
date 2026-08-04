! vybe-test: fortran/character_scan_verify_extended/ichar_uppercase_z_code
! origin: languages/fortran/tests/fortran/test_character_scan_verify_extended.rs
program t
if ((ichar('Z')) /= 90) then
    print *, "FAIL: want [90] got [", ichar('Z'), "]"
    stop 1
end if
end program t
