! vybe-test: fortran/character_scan_verify_extended/scan_forward_punctuation_comma
! origin: languages/fortran/tests/fortran/test_character_scan_verify_extended.rs
program t
character(len=10) :: s = 'value,more'
if ((scan(s, ',.;:')) /= 6) then
    print *, "FAIL: want [6] got [", scan(s, ',.;:'), "]"
    stop 1
end if
end program t
