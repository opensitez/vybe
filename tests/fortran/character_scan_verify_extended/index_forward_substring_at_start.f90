! vybe-test: fortran/character_scan_verify_extended/index_forward_substring_at_start
! origin: languages/fortran/tests/fortran/test_character_scan_verify_extended.rs
program t
character(len=7) :: s = 'fortran'
if ((index(s, 'for')) /= 1) then
    print *, "FAIL: want [1] got [", index(s, 'for'), "]"
    stop 1
end if
end program t
