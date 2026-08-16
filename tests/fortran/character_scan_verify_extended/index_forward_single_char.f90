! vybe-test: fortran/character_scan_verify_extended/index_forward_single_char
! origin: languages/fortran/tests/fortran/test_character_scan_verify_extended.rs
program t
character(len=11) :: s = 'mississippi'
if ((index(s, 's')) /= 3) then
    print *, "FAIL: want [3] got [", index(s, 's'), "]"
    stop 1
end if
end program t
