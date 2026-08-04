! vybe-test: fortran/character_scan_verify_extended/ichar_lowercase_a_code
! origin: languages/fortran/tests/fortran/test_character_scan_verify_extended.rs
program t
if ((ichar('a')) /= 97) then
    print *, "FAIL: want [97] got [", ichar('a'), "]"
    stop 1
end if
end program t
