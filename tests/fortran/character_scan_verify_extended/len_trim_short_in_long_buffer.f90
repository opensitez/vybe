! vybe-test: fortran/character_scan_verify_extended/len_trim_short_in_long_buffer
! origin: languages/fortran/tests/fortran/test_character_scan_verify_extended.rs
program t
character(len=10) :: s = 'go        '
if ((len_trim(s)) /= 2) then
    print *, "FAIL: want [2] got [", len_trim(s), "]"
    stop 1
end if
end program t
