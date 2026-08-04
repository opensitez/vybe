! vybe-test: fortran/character_scan_verify_extended/len_trim_full_no_trailing
! origin: languages/fortran/tests/fortran/test_character_scan_verify_extended.rs
program t
character(len=5) :: s = 'abcde'
if ((len_trim(s)) /= 5) then
    print *, "FAIL: want [5] got [", len_trim(s), "]"
    stop 1
end if
end program t
