! vybe-test: fortran/character_scan_verify_extended/scan_forward_no_match_returns_zero
! origin: languages/fortran/tests/fortran/test_character_scan_verify_extended.rs
program t
character(len=6) :: s = 'rhythm'
if ((scan(s, 'aeiou')) /= 0) then
    print *, "FAIL: want [0] got [", scan(s, 'aeiou'), "]"
    stop 1
end if
end program t
