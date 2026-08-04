! vybe-test: fortran/character_scan_verify_extended/len_trim_internal_spaces_counted
! origin: languages/fortran/tests/fortran/test_character_scan_verify_extended.rs
program t
character(len=7) :: s = 'a b c  '
if ((len_trim(s)) /= 5) then
    print *, "FAIL: want [5] got [", len_trim(s), "]"
    stop 1
end if
end program t
