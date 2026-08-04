! vybe-test: fortran/character_scan_verify_extended/ichar_digit_2
! origin: languages/fortran/tests/fortran/test_character_scan_verify_extended.rs
program t
if ((ichar('2')) /= 50) then
    print *, "FAIL: want [50] got [", ichar('2'), "]"
    stop 1
end if
end program t
