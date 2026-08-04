! vybe-test: fortran/character_scan_verify_extended/ichar_uppercase_m_code
! origin: languages/fortran/tests/fortran/test_character_scan_verify_extended.rs
program t
if ((ichar('M')) /= 77) then
    print *, "FAIL: want [77] got [", ichar('M'), "]"
    stop 1
end if
end program t
