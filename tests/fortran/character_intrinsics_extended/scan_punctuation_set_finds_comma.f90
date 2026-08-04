! vybe-test: fortran/character_intrinsics_extended/scan_punctuation_set_finds_comma
! origin: languages/fortran/tests/fortran/test_character_intrinsics_extended.rs
program t
character(len=11) :: s = 'value,more'
if ((scan(s, ',.;:')) /= 6) then
    print *, "FAIL: want [6] got [", scan(s, ',.;:'), "]"
    stop 1
end if
end program t
