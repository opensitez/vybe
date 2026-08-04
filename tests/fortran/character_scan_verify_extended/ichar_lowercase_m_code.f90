! vybe-test: fortran/character_scan_verify_extended/ichar_lowercase_m_code
! origin: languages/fortran/tests/fortran/test_character_scan_verify_extended.rs
program t
if ((ichar('m')) /= 109) then
    print *, "FAIL: want [109] got [", ichar('m'), "]"
    stop 1
end if
end program t
