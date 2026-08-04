! vybe-test: fortran/character_intrinsics_extended/scan_whitespace_set_in_mixed_text
! origin: languages/fortran/tests/fortran/test_character_intrinsics_extended.rs
program t
character(len=9) :: s = 'key=value'
if ((scan(s, ' =')) /= 4) then
    print *, "FAIL: want [4] got [", scan(s, ' ='), "]"
    stop 1
end if
end program t
