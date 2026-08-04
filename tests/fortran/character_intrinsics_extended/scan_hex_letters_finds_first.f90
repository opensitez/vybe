! vybe-test: fortran/character_intrinsics_extended/scan_hex_letters_finds_first
! origin: languages/fortran/tests/fortran/test_character_intrinsics_extended.rs
program t
character(len=8) :: s = '019af2b0'
if ((scan(s, 'abcdef')) /= 4) then
    print *, "FAIL: want [4] got [", scan(s, 'abcdef'), "]"
    stop 1
end if
end program t
