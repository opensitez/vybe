! vybe-test: fortran/character_scan_verify_extended/adjustl_preserves_internal_space
! origin: languages/fortran/tests/fortran/test_character_scan_verify_extended.rs
program t
character(len=10) :: s = '  ab cd   '
if (trim(trim(adjustl(s))) /= "ab cd") then
    print *, "FAIL: want [ab cd] got [", trim(adjustl(s)), "]"
    stop 1
end if
end program t
