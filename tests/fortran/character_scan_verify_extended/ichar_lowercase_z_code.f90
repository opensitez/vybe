! vybe-test: fortran/character_scan_verify_extended/ichar_lowercase_z_code
! origin: languages/fortran/tests/fortran/test_character_scan_verify_extended.rs
program t
if ((ichar('z')) /= 122) then
    print *, "FAIL: want [122] got [", ichar('z'), "]"
    stop 1
end if
end program t
