! vybe-test: fortran/character_scan_verify_extended/scan_forward_equals_in_pair
! origin: languages/fortran/tests/fortran/test_character_scan_verify_extended.rs
program t
character(len=9) :: s = 'key=value'
if ((scan(s, ' =')) /= 4) then
    print *, "FAIL: want [4] got [", scan(s, ' ='), "]"
    stop 1
end if
end program t
