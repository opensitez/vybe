! vybe-test: fortran/character_scan_verify_extended/len_trim_all_blanks_zero
! origin: languages/fortran/tests/fortran/test_character_scan_verify_extended.rs
program t
character(len=6) :: s = '      '
if ((len_trim(s)) /= 0) then
    print *, "FAIL: want [0] got [", len_trim(s), "]"
    stop 1
end if
end program t
