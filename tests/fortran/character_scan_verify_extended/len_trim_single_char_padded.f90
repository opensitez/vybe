! vybe-test: fortran/character_scan_verify_extended/len_trim_single_char_padded
! origin: languages/fortran/tests/fortran/test_character_scan_verify_extended.rs
program t
character(len=8) :: s = 'x       '
if ((len_trim(s)) /= 1) then
    print *, "FAIL: want [1] got [", len_trim(s), "]"
    stop 1
end if
end program t
