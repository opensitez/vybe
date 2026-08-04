! vybe-test: fortran/character_scan_verify_extended/verify_tab_not_in_letters
! origin: languages/fortran/tests/fortran/test_character_scan_verify_extended.rs
program t
character(len=3) :: s = 'a	b'
if ((verify(s, 'ab')) /= 2) then
    print *, "FAIL: want [2] got [", verify(s, 'ab'), "]"
    stop 1
end if
end program t
