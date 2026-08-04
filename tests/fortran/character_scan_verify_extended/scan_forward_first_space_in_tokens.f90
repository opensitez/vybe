! vybe-test: fortran/character_scan_verify_extended/scan_forward_first_space_in_tokens
! origin: languages/fortran/tests/fortran/test_character_scan_verify_extended.rs
program t
character(len=7) :: s = 'one two'
if ((scan(s, ' ')) /= 4) then
    print *, "FAIL: want [4] got [", scan(s, ' '), "]"
    stop 1
end if
end program t
