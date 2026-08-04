! vybe-test: fortran/character_scan_verify_extended/index_forward_substring_middle
! origin: languages/fortran/tests/fortran/test_character_scan_verify_extended.rs
program t
character(len=7) :: s = 'fortran'
if ((index(s, 'tra')) /= 3) then
    print *, "FAIL: want [3] got [", index(s, 'tra'), "]"
    stop 1
end if
end program t
