! vybe-test: fortran/character_scan_verify_extended/index_forward_overlap_pattern
! origin: languages/fortran/tests/fortran/test_character_scan_verify_extended.rs
program t
character(len=4) :: s = 'aaaa'
if ((index(s, 'aa')) /= 1) then
    print *, "FAIL: want [1] got [", index(s, 'aa'), "]"
    stop 1
end if
end program t
