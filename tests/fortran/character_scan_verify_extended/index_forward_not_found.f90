! vybe-test: fortran/character_scan_verify_extended/index_forward_not_found
! origin: languages/fortran/tests/fortran/test_character_scan_verify_extended.rs
program t
character(len=7) :: s = 'fortran'
if ((index(s, 'java')) /= 0) then
    print *, "FAIL: want [0] got [", index(s, 'java'), "]"
    stop 1
end if
end program t
